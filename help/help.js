function switchLang(lang) {
    const isEnglish = lang === 'en';
    document.querySelectorAll('.lang-en').forEach(el => {
        el.style.display = isEnglish ? 'block' : 'none';
    });
    document.querySelectorAll('.lang-ru').forEach(el => {
        el.style.display = isEnglish ? 'none' : 'block';
    });
    document.querySelectorAll('.lang-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.lang === lang);
    });
    document.documentElement.lang = lang;
}

document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.lang-btn').forEach(btn => {
        btn.addEventListener('click', () => switchLang(btn.dataset.lang));
    });
    switchLang('en');
});
