document.addEventListener('DOMContentLoaded', () => {
    // --- Element Selection ---
    const body = document.body;
    const initialView = document.getElementById('initial-view');
    const lisText = document.getElementById('lis-text');
    const lisSubtitle = document.getElementById('lis-subtitle');
    const mainTitle = document.getElementById('main-title');
    const mainContentArea = document.getElementById('main-content-area');
    const tickerLabel = document.getElementById('ticker-label');
    const tickerInputField = document.getElementById('ticker-input-field');
    const analyzeButton = document.getElementById('analyze-button');
    const pageHeader = document.getElementById('page-header');
    const authModal = document.getElementById('auth-modal');
    const closeModalButton = document.getElementById('close-modal-button');
    const scrollSpacer = document.getElementById('scroll-spacer');

    // --- Initial Checks & Setup ---
    if (!initialView || !lisText || !lisSubtitle) {
        console.error("Crucial initial elements (initialView, lisText, lisSubtitle) are missing!");
        return;
    }
    if (!mainContentArea || !mainTitle || !tickerLabel || !tickerInputField || !analyzeButton) {
        console.error("Crucial main content elements (mainContentArea, mainTitle, or ticker elements) are missing!");
        return;
    }
    if (!scrollSpacer || scrollSpacer.offsetHeight === 0) {
        console.warn("Scroll spacer not found or has no height. Scrolling animations may not work.");
        if(document.body.scrollHeight <= window.innerHeight) {
            document.body.style.minHeight = '150vh'; // Fallback
        }
    }

    // Set initial states using GSAP's autoAlpha. CSS should define the very initial state.
    gsap.set([initialView, lisText, lisSubtitle], { autoAlpha: 1 }); // Ensure visible from start
    gsap.set(mainContentArea, { autoAlpha: 0 }); // Ensure hidden from start
    gsap.set([mainTitle, tickerLabel, tickerInputField, analyzeButton], { autoAlpha: 0 }); // Ensure hidden from start

    // const isAuthenticated = document.querySelector('#main-nav a[href*="dashboard"]') !== null; // Removed as per new auth flow
    gsap.registerPlugin(ScrollTrigger);

    const expandingIElement = document.createElement('div');
    expandingIElement.setAttribute('id', 'expanding-i-proxy');
    expandingIElement.style.position = 'fixed';
    expandingIElement.style.top = '50%';
    expandingIElement.style.left = '50%';
    expandingIElement.style.width = 'clamp(10px, 3vw, 30px)';
    expandingIElement.style.height = 'clamp(50px, 20vh, 150px)';
    expandingIElement.style.backgroundColor = 'var(--bordeaux)';
    expandingIElement.style.transform = 'translate(-50%, -50%) scale(0)';
    expandingIElement.style.zIndex = '25';
    expandingIElement.style.opacity = '0';
    document.body.appendChild(expandingIElement);

    const tl = gsap.timeline({
        scrollTrigger: {
            trigger: body,
            start: "top top",
            end: "60% bottom",
            scrub: 1.5,
            // markers: true, // Uncomment for debugging scroll trigger points
            onUpdate: self => {
                if (self.progress > 0.35) {
                    pageHeader.classList.add('scrolled');
                } else {
                    pageHeader.classList.remove('scrolled');
                }
            }
        }
    });

    // --- Animation Sequence ---
    // 1. Fade out LIS text and subtitle
    tl.to([lisText, lisSubtitle], {
        autoAlpha: 0,
        y: "-15vh", // Less movement
        ease: "power1.in",
        duration: 0.3 // Faster fade out
    }, 0);

    // 2. Start scaling and fading in the 'expandingIElement'
    tl.to(expandingIElement, {
        opacity: 1,
        scale: 250,
        ease: "power2.inOut",
        duration: 0.4 // Slightly faster expansion
    }, 0.1);

    // 3. Fade out initialView's own white background to transparent
    tl.to(initialView, {
        backgroundColor: 'rgba(255, 255, 255, 0)',
        duration: 0.3, // Faster fade
        ease: "power1.out"
    }, 0.15);

    // 4. Change body background to solid Bordeaux
    tl.to(body, {
        backgroundColor: 'var(--bordeaux)',
        duration: 0.01,
        ease: "none"
    }, 0.45); // Sooner

    // 5. Hide initialView and proxy 'I' after transition.
    tl.to(initialView, { autoAlpha: 0, duration: 0.01 }, 0.5) // Sooner
      .to(expandingIElement, { opacity: 0, duration: 0.01, onComplete: () => expandingIElement.style.display = 'none' }, 0.5); // Sooner

    // 6. Fade in the main content area
    tl.to(mainContentArea, {
        autoAlpha: 1,
        duration: 0.3, // Faster
        ease: "power2.out"
    }, 0.55); // Sooner

    // 7. Fade in the "Lisbon Investment Society" title at the top
    tl.to(mainTitle, {
        autoAlpha: 1,
        y: 0,
        duration: 0.3, // Faster
        ease: "power1.out"
    }, 0.65); // Sooner

    // 8. Fade in Ticker elements sequentially
    tl.to(tickerLabel, { autoAlpha: 1, duration: 0.3, ease: "power1.out" }, 0.75) // Faster and sooner
      .to(tickerInputField, { autoAlpha: 1, duration: 0.3, ease: "power1.out" }, 0.8) // Faster and sooner
      .to(analyzeButton, { autoAlpha: 1, duration: 0.3, ease: "power1.out" }, 0.85); // Faster and sooner

    // --- Modal Logic ---
    function showAuthModal() {
        if (authModal) authModal.classList.remove('hidden');
    }
    function hideAuthModal() {
        if (authModal) authModal.classList.add('hidden');
    }

    // --- Analysis Submission Logic ---
    function submitTickerForAnalysis(tickerValue) {
        if (!tickerValue || tickerValue.trim() === '') {
            alert('Ticker symbol cannot be empty for submission.');
            return;
        }
        const form = document.createElement('form');
        form.method = 'POST';
        form.action = '/analyze';
        const hiddenField = document.createElement('input');
        hiddenField.type = 'hidden';
        hiddenField.name = 'ticker';
        hiddenField.value = tickerValue.trim();
        form.appendChild(hiddenField);
        document.body.appendChild(form);
        form.submit();
    }

    if (analyzeButton) {
        analyzeButton.addEventListener('click', (e) => {
            e.preventDefault(); // Prevent default form submission

            const tickerValue = tickerInputField.value.trim();
            if (tickerValue === '') {
                alert('Please enter a ticker symbol.');
                return;
            }

            fetch('/check_auth')
                .then(response => response.json())
                .then(data => {
                    if (data.authenticated) {
                        submitTickerForAnalysis(tickerValue);
                    } else {
                        // User is not authenticated, show modal and store ticker
                        showAuthModal();
                        localStorage.setItem('pendingTicker', tickerValue);
                    }
                })
                .catch(error => {
                    console.error('Error checking authentication:', error);
                    // Optionally, show a generic error message to the user or fallback to showing the modal
                });
        });
    }

    if (closeModalButton) closeModalButton.addEventListener('click', hideAuthModal);
    if (authModal) {
        authModal.addEventListener('click', (e) => {
            if (e.target === authModal) hideAuthModal();
        });
    }

    // --- Handle Pending Ticker on Page Load ---
    const pendingTicker = localStorage.getItem('pendingTicker');
    if (pendingTicker) {
        if (tickerInputField) {
            tickerInputField.value = pendingTicker;
        }
        localStorage.removeItem('pendingTicker');
    }

    // --- Reduced motion fallback ---
    const mediaQueryReducedMotion = window.matchMedia('(prefers-reduced-motion: reduce)');
    if (mediaQueryReducedMotion.matches) {
        if (tl) tl.kill();
        if (initialView) gsap.set(initialView, { autoAlpha: 0, display: 'none' });
        if (expandingIElement) expandingIElement.style.display = 'none';
        body.style.backgroundColor = 'var(--bordeaux)';
        if (pageHeader) pageHeader.classList.add('scrolled');
        gsap.set(mainContentArea, { autoAlpha: 1 });
        gsap.set([mainTitle, tickerLabel, tickerInputField, analyzeButton], { autoAlpha: 1 });
    }
});
