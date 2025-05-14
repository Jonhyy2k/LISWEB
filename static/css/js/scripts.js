document.addEventListener('DOMContentLoaded', () => {
    const landing = document.getElementById('landing');
    const mainPage = document.getElementById('main-page');
    const lisText = document.getElementById('lis-text');
    const lisFull = document.getElementById('lis-full');
    const lisTitle = document.getElementById('lis-title');
    const tickerForm = document.getElementById('ticker-form') || document.getElementById('auth-links');

    // GSAP Animation
    let animationTriggered = false;
    window.addEventListener('scroll', () => {
        if (window.scrollY > 50 && !animationTriggered) {
            animationTriggered = true;
            gsap.to(landing, { opacity: 0, duration: 1, ease: 'power2.out' });
            gsap.to(lisText, { 
                scale: 10, 
                y: '-50vh', 
                duration: 1.5, 
                ease: 'power2.inOut', 
                onComplete: () => {
                    landing.style.display = 'none';
                    mainPage.classList.remove('hidden');
                    gsap.to(mainPage, { opacity: 1, duration: 1 });
                    gsap.to(lisTitle, { opacity: 1, y: 0, duration: 1, delay: 0.5 });
                    gsap.to(tickerForm, { opacity: 1, y: 0, duration: 1, delay: 1 });
                }
            });
        }
    });
});