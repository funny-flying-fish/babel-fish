(() => {
    function resolveTabName(tabName, aliases) {
        return aliases[tabName] || tabName;
    }

    function getTabButton(tabName) {
        return document.querySelector(`.tab-btn[data-tab="${tabName}"]`);
    }

    const UTILITY_TAB_IDS = new Set(['xlsx-to-lokalise', 'find-diff', 'merge-xlsx']);

    function isUtilitiesPage() {
        return /utilities\.html$/i.test(location.pathname);
    }

    function redirectToUtilitiesTab(tabName) {
        if (!tabName || isUtilitiesPage() || !UTILITY_TAB_IDS.has(tabName)) return false;
        window.location.replace(`utilities.html#${tabName}`);
        return true;
    }

    function initBabelFishNav(options = {}) {
        const {
            defaultSection = 'print',
            defaultTab = '',
            pageSection = defaultSection,
            tabTitles = {},
            tabHashAliases = {}
        } = options;

        const sectionNav = document.querySelector('.section-nav');
        const tabButtons = Array.from(document.querySelectorAll('.tab-btn[data-tab]'));
        if (!sectionNav || !tabButtons.length) return;

        let activeSection = defaultSection;

        function tabsInSection(section) {
            return tabButtons.filter(btn => btn.dataset.section === section);
        }

        function defaultTabForSection(section) {
            const tabs = tabsInSection(section);
            return tabs[0] ? tabs[0].dataset.tab : '';
        }

        function updateSectionNav() {
            sectionNav.querySelectorAll('.section-nav-btn').forEach((btn) => {
                const isActive = btn.dataset.section === activeSection && btn.tagName === 'BUTTON';
                btn.classList.toggle('active', isActive);
            });
        }

        function updateTabVisibility() {
            tabButtons.forEach((btn) => {
                const visible = btn.dataset.section === activeSection;
                btn.hidden = !visible;
                btn.style.display = visible ? '' : 'none';
            });
        }

        function updateLogTitle(tabName) {
            const logTitle = document.getElementById('logTitle');
            if (logTitle && tabTitles[tabName]) {
                logTitle.textContent = tabTitles[tabName];
            }
        }

        function activateTab(tabName, opts = {}) {
            const resolved = resolveTabName(tabName, tabHashAliases);
            const btn = getTabButton(resolved);
            if (!btn || btn.hidden) return false;

            if (!opts.skipSection && btn.dataset.section) {
                activeSection = btn.dataset.section;
                updateSectionNav();
                updateTabVisibility();
            }

            tabButtons.forEach(b => b.classList.remove('active'));
            document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
            btn.classList.add('active');
            const tabEl = document.getElementById('tab-' + resolved);
            if (tabEl) tabEl.classList.add('active');
            updateLogTitle(resolved);
            return resolved;
        }

        function activateSection(section, opts = {}) {
            if (!tabsInSection(section).length) return;
            activeSection = section;
            updateSectionNav();
            updateTabVisibility();
            if (!opts.keepTab) {
                const tab = opts.tab || defaultTabForSection(section);
                if (tab) activateTab(tab, { skipSection: true });
            }
        }

        sectionNav.querySelectorAll('.section-nav-btn').forEach((btn) => {
            btn.addEventListener('click', (event) => {
                if (btn.tagName === 'A') return;
                event.preventDefault();
                activateSection(btn.dataset.section);
                const tab = defaultTabForSection(btn.dataset.section);
                if (tab) {
                    history.replaceState(null, '', '#' + tab);
                }
            });
        });

        tabButtons.forEach((btn) => {
            btn.addEventListener('click', () => {
                const tabName = btn.dataset.tab;
                const resolved = activateTab(tabName);
                if (resolved) history.replaceState(null, '', '#' + resolved);
            });
        });

        const hash = window.location.hash.slice(1);
        const resolvedHash = resolveTabName(hash, tabHashAliases);

        if (hash && redirectToUtilitiesTab(resolvedHash)) {
            return;
        }

        if (hash && getTabButton(resolvedHash)) {
            const btn = getTabButton(resolvedHash);
            activeSection = btn.dataset.section || defaultSection;
            updateSectionNav();
            updateTabVisibility();
            activateTab(resolvedHash, { skipSection: true });
        } else {
            activateSection(defaultSection, { tab: defaultTab || defaultTabForSection(defaultSection) });
        }

        window.addEventListener('hashchange', () => {
            const newHash = window.location.hash.slice(1);
            if (!newHash) return;
            const resolved = resolveTabName(newHash, tabHashAliases);
            if (redirectToUtilitiesTab(resolved)) return;
            if (getTabButton(resolved)) {
                activateTab(resolved);
            }
        });
    }

    window.initBabelFishNav = initBabelFishNav;
})();
