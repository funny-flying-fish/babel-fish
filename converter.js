const NBSP = "\u00A0";
const DEFAULT_SETTINGS = {
    numericFormatting: true,
    numberSeparators: true,
    percentFormatting: true,
    currencyFormatting: true,
    unitSpacing: true,
    unitCaseNormalization: true,
    stripOuterQuotes: true,
    globalRules: true,
    languageRules: true
};

function getSettings() {
    const current = window.nbspSettings || {};
    return { ...DEFAULT_SETTINGS, ...current };
}

function getTxtSettings() {
    const current = window.nbspSettingsTxt || {};
    return { ...DEFAULT_SETTINGS, ...current };
}

(() => {
    const NUMBER_PATTERN = '(?:\\d[\\d.,\\u00A0 ]*\\d|\\d)';
    const UNIT_TOKENS = [
        'W', 'kW', 'MW', 'GW', 'hp', 'PS', 'CV',
        'V', 'kV', 'mV', 'A', 'mA', 'Hz', 'kHz', 'MHz', 'GHz', 'Ohm', 'Ω',
        'Wh', 'kWh', 'MWh', 'J', 'kJ', 'F', 'µF', 'μF', 'nF',
        'N', 'kN', 'Kgf', 'kgf', 'Pa', 'kPa', 'bar', 'Nm',
        'km', 'm', 'cm', 'mm', 'kg', 'g', 'mg', 'l', 'ml'
    ];

    function logChange(logger, rule, before, after) {
        if (!logger || before === after) return;
        logger(rule, before, after);
    }

    function applyGlobalRules(text, enabled, logger) {
        if (!enabled) return text;
        let result = text;

        // Initials and surname
        result = result.replace(/\b([A-ZÀ-ÖØ-Þ])\.\s+([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'-]*)/g, (match, initial, surname) => {
            const updated = `${initial}.${NBSP}${surname}`;
            logChange(logger, 'NBSP after initials', match, updated);
            return updated;
        });

        // Abbreviations and numerals
        result = result.replace(/\b(p\.|№|Vol\.)\s+(\d+)/gi, (match, abbr, number) => {
            const updated = `${abbr}${NBSP}${number}`;
            logChange(logger, 'NBSP after abbreviations', match, updated);
            return updated;
        });

        // Dates: "Jan 15, 2026 year"
        result = result.replace(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2},)/g, (match, month, day) => {
            const updated = `${month}${NBSP}${day}`;
            logChange(logger, 'NBSP in dates', match, updated);
            return updated;
        });
        result = result.replace(/\b(\d{4})\s+(year)\b/gi, (match, year, word) => {
            const updated = `${year}${NBSP}${word}`;
            logChange(logger, 'NBSP in dates', match, updated);
            return updated;
        });

        return result;
    }

    function applyUnitSpacing(text, enabled, logger) {
        if (!enabled) return text;

        const simpleUnits = `(?:${UNIT_TOKENS.join('|')})`;
        const compositeUnits = '(?:[A-Za-zΩµμ]+[0-9]*(?:[./][A-Za-zΩµμ0-9]+)+)';
        const spacedUnitsRegex = new RegExp(`(${NUMBER_PATTERN})\\s+(${simpleUnits}|${compositeUnits})\\b`, 'gi');
        const tightUnitsRegex = new RegExp(`(${NUMBER_PATTERN})(${simpleUnits}|${compositeUnits})\\b`, 'gi');

        let result = text.replace(spacedUnitsRegex, (match, number, unit) => {
            const updated = `${number}${NBSP}${unit}`;
            logChange(logger, 'NBSP before units', match, updated);
            return updated;
        });
        result = result.replace(tightUnitsRegex, (match, number, unit) => {
            const updated = `${number}${NBSP}${unit}`;
            logChange(logger, 'NBSP before units', match, updated);
            return updated;
        });
        return result;
    }

    function applyUnitCaseNormalization(text, enabled, logger) {
        if (!enabled) return text;

        const normalizedMap = {
            w: 'W',
            kw: 'kW',
            mw: 'MW',
            gw: 'GW',
            hp: 'hp',
            ps: 'PS',
            cv: 'CV',
            v: 'V',
            kv: 'kV',
            mv: 'mV',
            a: 'A',
            ma: 'mA',
            hz: 'Hz',
            khz: 'kHz',
            mhz: 'MHz',
            ghz: 'GHz',
            ohm: 'Ohm',
            wh: 'Wh',
            kwh: 'kWh',
            mwh: 'MWh',
            j: 'J',
            kj: 'kJ',
            f: 'F',
            uf: 'μF',
            µf: 'μF',
            μf: 'μF',
            nf: 'nF',
            n: 'N',
            kn: 'kN',
            kgf: 'kgf',
            pa: 'Pa',
            kpa: 'kPa',
            bar: 'bar',
            km: 'km',
            m: 'm',
            cm: 'cm',
            mm: 'mm',
            kg: 'kg',
            g: 'g',
            mg: 'mg',
            l: 'l',
            ml: 'ml',
            nm: 'Nm'
        };

        function normalizeSimpleUnit(token) {
            if (token.includes('Ω')) return 'Ω';
            const lower = token.toLowerCase();
            return normalizedMap[lower] || token;
        }

        function normalizeCompositeUnit(token) {
            const lower = token.toLowerCase();
            if (lower === 'kg/h/n') return 'kg/h/N';
            if (lower === 'm/s2') return 'm/s²';

            return token.replace(/[A-Za-zΩµμ]+[0-9]*/g, (segment) => {
                const match = segment.match(/^([A-Za-zΩµμ]+)(\d*)$/);
                if (!match) return segment;
                const letters = normalizeSimpleUnit(match[1]);
                return `${letters}${match[2]}`;
            });
        }

        const simpleUnits = `(?:${UNIT_TOKENS.join('|')})`;
        const compositeUnits = '(?:[A-Za-zΩµμ]+[0-9]*(?:[./][A-Za-zΩµμ0-9]+)+)';
        const unitRegex = new RegExp(`(${NUMBER_PATTERN})([\\u00A0\\s]+)(${simpleUnits}|${compositeUnits})\\b`, 'gi');

        return text.replace(unitRegex, (match, num, sep, unit) => {
            const normalizedUnit = /[./]/.test(unit)
                ? normalizeCompositeUnit(unit)
                : normalizeSimpleUnit(unit);
            const updated = `${num}${sep}${normalizedUnit}`;
            logChange(logger, 'Unit case normalization', match, updated);
            return updated;
        });
    }

    function applyLanguageRules(text, langCode, enabled, logger) {
        if (!enabled) return text;
        const lang = (langCode || "").toUpperCase();
        let result = text;

        if (lang === "EN") {
            result = result.replace(/\b(a|an|the|I)\s+/gi, (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'EN NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "FR") {
            result = result.replace(/\b(à|y|en|le|la|les|un|une|de|du)\s+/gi, (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'FR NBSP after short words', match, updated);
                return updated;
            });
            result = result.replace(/([^\s])\s*([;:!?])/g, (match, left, punctuation) => {
                const updated = `${left}${NBSP}${punctuation}`;
                logChange(logger, 'FR NBSP before punctuation', match, updated);
                return updated;
            });
            result = result.replace(/«[\u00A0\u202F]+/g, (match) => {
                const updated = '«';
                logChange(logger, 'FR guillemet spacing removed', match, updated);
                return updated;
            });
            result = result.replace(/[\u00A0\u202F]+»/g, (match) => {
                const updated = '»';
                logChange(logger, 'FR guillemet spacing removed', match, updated);
                return updated;
            });
        } else if (lang === "ES") {
            result = result.replace(/\b(a|e|i|o|u|y|la|el|un)\s+/gi, (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'ES NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "DE") {
            result = result.replace(/\b(der|die|das|ein|eine|Dr\.|Prof\.|Hr\.|Fr\.)\s+/gi, (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'DE NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "NL") {
            result = result.replace(/\b(de|het|een|in|op|te|ten|ter)\s+/gi, (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'NL NBSP after short words', match, updated);
                return updated;
            });
        }

        return result;
    }

    function applyNumberSeparators(text, langCode, enabled, logger) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const thousandsSep = lang === "FR" ? NBSP : (lang === "EN" ? "," : ".");
        const decimalSep = lang === "EN" ? "." : ",";

        function parseNumberString(numStr) {
            const cleaned = numStr.replace(/\u00A0/g, ' ');
            const sepMatches = cleaned.match(/[.,]/g);
            const sepCount = sepMatches ? sepMatches.length : 0;
            const lastDot = cleaned.lastIndexOf('.');
            const lastComma = cleaned.lastIndexOf(',');
            const decimalIndex = Math.max(lastDot, lastComma);

            let intPart = cleaned;
            let fracPart = '';

            if (decimalIndex > -1) {
                const after = cleaned.slice(decimalIndex + 1);
                const before = cleaned.slice(0, decimalIndex);
                const digitsAfter = after.match(/^\d+$/) ? after.length : 0;

                if (!(sepCount === 1 && digitsAfter === 3)) {
                    intPart = before;
                    fracPart = after;
                }
            }

            intPart = intPart.replace(/[.,\s]/g, '');
            if (!intPart) return null;
            if (fracPart) {
                fracPart = fracPart.replace(/[.,\s]/g, '');
            }
            return { intPart, fracPart };
        }

        function formatIntegerPart(intPart, separator, lang) {
            if (lang === "ES" && intPart.length <= 4) {
                return intPart;
            }
            return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, separator);
        }

        const numberRegex = new RegExp(NUMBER_PATTERN, 'g');
        return text.replace(numberRegex, (match) => {
            const parts = parseNumberString(match);
            if (!parts) return match;
            const formattedInt = formatIntegerPart(parts.intPart, thousandsSep, lang);
            const updated = parts.fracPart ? `${formattedInt}${decimalSep}${parts.fracPart}` : formattedInt;
            logChange(logger, 'Number separators', match, updated);
            return updated;
        });
    }

    function formatPercent(text, langCode, enabled, logger) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const numberRegex = new RegExp(`(${NUMBER_PATTERN})\\s*%`, 'g');
        if (lang === "EN") {
            return text.replace(numberRegex, (match, number) => {
                const updated = `${number}%`;
                logChange(logger, 'Percent formatting', match, updated);
                return updated;
            });
        }
        return text.replace(numberRegex, (match, number) => {
            const updated = `${number}${NBSP}%`;
            logChange(logger, 'Percent formatting', match, updated);
            return updated;
        });
    }

    function formatCurrency(text, langCode, enabled, logger) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const numberCapture = NUMBER_PATTERN;
        const prefixRegex = new RegExp(`([€$£])\\s*(${numberCapture})`, 'g');
        const suffixRegex = new RegExp(`(${numberCapture})\\s*([€$£])`, 'g');

        if (lang === "EN") {
            let result = text.replace(suffixRegex, (match, number, symbol) => {
                const updated = `${symbol}${number}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            result = result.replace(prefixRegex, (match, symbol, number) => {
                const updated = `${symbol}${number}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            return result;
        }

        if (lang === "NL") {
            let result = text.replace(suffixRegex, (match, number, symbol) => {
                const updated = `${symbol}${NBSP}${number}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            result = result.replace(prefixRegex, (match, symbol, number) => {
                const updated = `${symbol}${NBSP}${number}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            return result;
        }

        if (lang === "FR" || lang === "DE" || lang === "ES") {
            let result = text.replace(prefixRegex, (match, symbol, number) => {
                const updated = `${number}${NBSP}${symbol}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            result = result.replace(suffixRegex, (match, number, symbol) => {
                const updated = `${number}${NBSP}${symbol}`;
                logChange(logger, 'Currency formatting', match, updated);
                return updated;
            });
            return result;
        }

        return text;
    }

    function applyNumericFormatting(text, langCode, enabled, logger) {
        if (!enabled) return text;
        const settings = getSettings();
        let result = applyNumberSeparators(text, langCode, settings.numberSeparators, logger);
        result = formatPercent(result, langCode, settings.percentFormatting, logger);
        result = formatCurrency(result, langCode, settings.currencyFormatting, logger);
        return result;
    }

    window.applyNbspRules = function applyNbspRules(value, langCode, options = {}) {
        if (value === null || value === undefined) return '';
        const text = String(value);
        if (!text) return text;

        const settings = options.settings
            ? { ...DEFAULT_SETTINGS, ...options.settings }
            : getSettings();
        const logger = typeof options.logger === 'function' ? options.logger : null;
        const withNumeric = applyNumericFormatting(text, langCode, settings.numericFormatting, logger);
        const withUnits = applyUnitSpacing(withNumeric, settings.unitSpacing, logger);
        const withUnitCase = applyUnitCaseNormalization(withUnits, settings.unitCaseNormalization, logger);
        const withGlobal = applyGlobalRules(withUnitCase, settings.globalRules, logger);
        return applyLanguageRules(withGlobal, langCode, settings.languageRules, logger);
    };
})();

// File references
let xlsxFile = null;
let txtFile = null;

// Elements
const xlsxInput = document.getElementById('xlsxInput');
const txtInput = document.getElementById('txtInput');
const xlsxToTxtBtn = document.getElementById('xlsxToTxtBtn');
const txtToXlsxBtn = document.getElementById('txtToXlsxBtn');
const txtDownload = document.getElementById('txtDownload');
const xlsxDownload = document.getElementById('xlsxDownload');
const xlsxError = document.getElementById('xlsxError');
const txtError = document.getElementById('txtError');
const xlsxOutputEncoding = document.getElementById('xlsxOutputEncoding');
const txtInputEncoding = document.getElementById('txtInputEncoding');
const ruleNumeric = document.getElementById('ruleNumeric');
const ruleUnitSpacing = document.getElementById('ruleUnitSpacing');
const ruleUnitCase = document.getElementById('ruleUnitCase');
const ruleStripQuotes = document.getElementById('ruleStripQuotes');
const ruleGlobal = document.getElementById('ruleGlobal');
const ruleLanguage = document.getElementById('ruleLanguage');
const ruleNumericTxt = document.getElementById('ruleNumericTxt');
const ruleUnitSpacingTxt = document.getElementById('ruleUnitSpacingTxt');
const ruleUnitCaseTxt = document.getElementById('ruleUnitCaseTxt');
const ruleStripQuotesTxt = document.getElementById('ruleStripQuotesTxt');
const ruleGlobalTxt = document.getElementById('ruleGlobalTxt');
const ruleLanguageTxt = document.getElementById('ruleLanguageTxt');
const openRules = document.getElementById('openRules');
const rulesModal = document.getElementById('rulesModal');
const closeRules = document.getElementById('closeRules');
const rulesContent = document.getElementById('rulesContent');
const logTitle = document.getElementById('logTitle');
const logCount = document.getElementById('logCount');
const logEmpty = document.getElementById('logEmpty');
const logList = document.getElementById('logList');

window.nbspSettings = {
    numericFormatting: ruleNumeric.checked,
    unitSpacing: ruleUnitSpacing.checked,
    unitCaseNormalization: ruleUnitCase.checked,
    stripOuterQuotes: ruleStripQuotes.checked,
    globalRules: ruleGlobal.checked,
    languageRules: ruleLanguage.checked
};

window.nbspSettingsTxt = {
    numericFormatting: ruleNumericTxt.checked,
    unitSpacing: ruleUnitSpacingTxt.checked,
    unitCaseNormalization: ruleUnitCaseTxt.checked,
    stripOuterQuotes: ruleStripQuotesTxt.checked,
    globalRules: ruleGlobalTxt.checked,
    languageRules: ruleLanguageTxt.checked
};

function syncNbspSettings() {
    window.nbspSettings = {
        numericFormatting: ruleNumeric.checked,
        unitSpacing: ruleUnitSpacing.checked,
        unitCaseNormalization: ruleUnitCase.checked,
        stripOuterQuotes: ruleStripQuotes.checked,
        globalRules: ruleGlobal.checked,
        languageRules: ruleLanguage.checked
    };
}

function syncNbspSettingsTxt() {
    window.nbspSettingsTxt = {
        numericFormatting: ruleNumericTxt.checked,
        unitSpacing: ruleUnitSpacingTxt.checked,
        unitCaseNormalization: ruleUnitCaseTxt.checked,
        stripOuterQuotes: ruleStripQuotesTxt.checked,
        globalRules: ruleGlobalTxt.checked,
        languageRules: ruleLanguageTxt.checked
    };
}

ruleNumeric.addEventListener('change', syncNbspSettings);
ruleUnitSpacing.addEventListener('change', syncNbspSettings);
ruleUnitCase.addEventListener('change', syncNbspSettings);
ruleStripQuotes.addEventListener('change', syncNbspSettings);
ruleGlobal.addEventListener('change', syncNbspSettings);
ruleLanguage.addEventListener('change', syncNbspSettings);

ruleNumericTxt.addEventListener('change', syncNbspSettingsTxt);
ruleUnitSpacingTxt.addEventListener('change', syncNbspSettingsTxt);
ruleUnitCaseTxt.addEventListener('change', syncNbspSettingsTxt);
ruleStripQuotesTxt.addEventListener('change', syncNbspSettingsTxt);
ruleGlobalTxt.addEventListener('change', syncNbspSettingsTxt);
ruleLanguageTxt.addEventListener('change', syncNbspSettingsTxt);

async function openRulesModal() {
    rulesModal.classList.add('is-open');
    rulesModal.setAttribute('aria-hidden', 'false');
    try {
        const response = await fetch('rules.html', { cache: 'no-store' });
        if (!response.ok) throw new Error('Failed to load rules');
        const text = await response.text();
        rulesContent.innerHTML = text;
    } catch (err) {
        rulesContent.textContent = 'Unable to load rules.';
    }
}

function closeRulesModal() {
    rulesModal.classList.remove('is-open');
    rulesModal.setAttribute('aria-hidden', 'true');
}

openRules.addEventListener('click', openRulesModal);
closeRules.addEventListener('click', closeRulesModal);
rulesModal.addEventListener('click', (e) => {
    if (e.target === rulesModal) closeRulesModal();
});
document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && rulesModal.classList.contains('is-open')) {
        closeRulesModal();
    }
});

let logEntries = [];

function addLog(rule, before, after, context) {
    logEntries.push({ rule, before, after, context });
}

function columnToLetters(index) {
    let result = '';
    let current = index;
    while (current > 0) {
        const remainder = (current - 1) % 26;
        result = String.fromCharCode(65 + remainder) + result;
        current = Math.floor((current - 1) / 26);
    }
    return result;
}

function escapeHtml(value) {
    return String(value)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function formatWhitespaceValue(value) {
    const text = String(value);
    const changed = new Set();

    let start = 0;
    while (start < text.length && text[start] === ' ') {
        changed.add(start);
        start += 1;
    }

    let end = text.length - 1;
    while (end >= 0 && text[end] === ' ') {
        changed.add(end);
        end -= 1;
    }

    let i = 0;
    while (i < text.length) {
        if (text[i] !== ' ') {
            i += 1;
            continue;
        }
        const runStart = i;
        while (i < text.length && text[i] === ' ') {
            i += 1;
        }
        const runLength = i - runStart;
        if (runLength >= 2) {
            for (let j = runStart; j < runStart + runLength; j++) {
                changed.add(j);
            }
        }
    }

    let result = '';
    for (let idx = 0; idx < text.length; idx++) {
        const char = text[idx];
        if (changed.has(idx) && char === ' ') {
            result += '<span class="log-ws">□</span>';
        } else if (char === '\r' || char === '\n') {
            result += '<span class="log-ws">□</span>';
        } else {
            result += escapeHtml(char);
        }
    }
    return result;
}

function formatLineBreakValue(value) {
    const text = String(value);
    let result = '';
    for (const char of text) {
        if (char === '\r' || char === '\n') {
            result += '<span class="log-ws">□</span>';
        } else {
            result += escapeHtml(char);
        }
    }
    return result;
}

function formatCellRefHtml(cellRef) {
    if (!cellRef) return '';
    const match = cellRef.match(/^([A-Z]+)(\d+)$/);
    if (!match) return escapeHtml(cellRef);
    return `${match[1]}<strong>${match[2]}</strong>`;
}

function renderLog() {
    logList.innerHTML = '';
    if (logEntries.length === 0) {
        logEmpty.style.display = 'block';
        logCount.textContent = '0 changes';
        return;
    }
    logEmpty.style.display = 'none';
    logCount.textContent = `${logEntries.length} changes`;
    logEntries.forEach((entry) => {
        const item = document.createElement('li');
        const contextParts = [];
        if (entry.context && entry.context.cellRef) {
            contextParts.push(formatCellRefHtml(entry.context.cellRef));
        }
        const contextHtml = contextParts.length ? `${contextParts.join(' ')} ` : '';
        if (entry.rule === 'Whitespace normalization' || entry.rule === 'Line breaks removed') {
            const valueHtml = entry.rule === 'Line breaks removed'
                ? formatLineBreakValue(entry.before)
                : formatWhitespaceValue(entry.before);
            item.innerHTML = `${contextHtml}<span class="log-rule">${escapeHtml(entry.rule)}</span>: ` +
                `<span class="log-line">"${valueHtml}"</span>`;
        } else {
            const beforeHtml = escapeHtml(entry.before);
            const afterHtml = escapeHtml(entry.after);
            item.innerHTML = `${contextHtml}<span class="log-rule">${escapeHtml(entry.rule)}</span>: ` +
                `<span class="log-before">"${beforeHtml}"</span> → ` +
                `<span class="log-after">"${afterHtml}"</span>`;
        }
        logList.appendChild(item);
    });
}

function resetLog(title) {
    logEntries = [];
    logTitle.textContent = title;
    renderLog();
}

// Event listeners
xlsxInput.addEventListener('change', (e) => {
    xlsxFile = e.target.files[0];
    xlsxToTxtBtn.disabled = !xlsxFile;
    txtDownload.innerHTML = '';
    xlsxError.textContent = '';
});

txtInput.addEventListener('change', (e) => {
    txtFile = e.target.files[0];
    txtToXlsxBtn.disabled = !txtFile;
    xlsxDownload.innerHTML = '';
    txtError.textContent = '';
});

xlsxToTxtBtn.addEventListener('click', convertXlsxToTxt);
txtToXlsxBtn.addEventListener('click', convertTxtToXlsx);

// Get current date in YY-MM-DD format
function getCurrentDate() {
    const now = new Date();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    const year = String(now.getFullYear()).slice(-2);
    return `${year}-${month}-${day}`;
}

// Convert locale code (en_EN) to 2-letter uppercase (EN) for TXT
function localeToShort(locale) {
    if (!locale || typeof locale !== 'string') return locale;
    // Handle format like "en_EN" -> "EN"
    const match = locale.match(/^([a-z]{2})_([A-Z]{2})$/i);
    if (match) {
        return match[2].toUpperCase();
    }
    return locale;
}

// Convert 2-letter code (EN) to locale format (en_EN) for XLSX
function shortToLocale(code) {
    if (!code || typeof code !== 'string') return code;
    // Handle format like "EN" -> "en_EN"
    const match = code.match(/^([A-Z]{2})$/i);
    if (match) {
        const lower = code.toLowerCase();
        const upper = code.toUpperCase();
        return `${lower}_${upper}`;
    }
    return code;
}

// Parse XLSX file and return 2D array
function parseXLSX(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });
                resolve(jsonData);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Parse TXT file with specified encoding and return 2D array
function parseTXT(file, encoding) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const buffer = e.target.result;
                const decoder = new TextDecoder(encoding);
                const text = decoder.decode(buffer);
                // Handle both \r\n (Windows) and \n (Unix) line endings
                const lines = text.split(/\r?\n/).filter(line => line.trim() !== '');
                const data = lines.map(line => line.split('\t'));
                resolve(data);
            } catch (err) {
                reject(err);
            }
        };
        reader.onerror = reject;
        reader.readAsArrayBuffer(file);
    });
}

// Transpose 2D array
function transpose(matrix) {
    if (matrix.length === 0) return [];
    const rows = matrix.length;
    const cols = Math.max(...matrix.map(row => row.length));
    const result = [];
    for (let c = 0; c < cols; c++) {
        result[c] = [];
        for (let r = 0; r < rows; r++) {
            result[c][r] = matrix[r][c] !== undefined ? matrix[r][c] : '';
        }
    }
    return result;
}

function stripOuterQuotes(value, enabled, logger, context) {
    if (value === null || value === undefined) return '';
    const text = String(value);
    if (!enabled) return text;
    if (text.length < 2) return text;
    const pairs = {
        '"': '"',
        '«': '»',
        '“': '”',
        '„': '“'
    };
    const first = text[0];
    const last = text[text.length - 1];
    if (pairs[first] && pairs[first] === last) {
        const stripped = text.slice(1, -1);
        if (logger && stripped !== text) {
            logger('Outer quotes removed', text, stripped, context);
        }
        return stripped;
    }
    return text;
}

function applyFrenchNarrowSpaces(text, enabled, logger, context) {
    if (!enabled) return text;
    let result = text;
    const normalized = result.replace(/\u202F/g, NBSP);
    if (logger && normalized !== result) {
        logger('FR narrow spaces normalized', result, normalized, context);
    }
    result = normalized;
    const beforeQuotes = result;
    result = result
        .replace(/«[\u00A0\u202F]+/g, '«')
        .replace(/[\u00A0\u202F]+»/g, '»');
    if (logger && result !== beforeQuotes) {
        logger('FR guillemet spacing removed', beforeQuotes, result, context);
    }
    return result;
}

function normalizeTxtCell(value, logger, context, langCode) {
    if (value === null || value === undefined) return '';
    let str = String(value);
    const settings = getSettings();
    const lang = (langCode || "").toUpperCase();
    if (lang === "FR") {
        str = applyFrenchNarrowSpaces(str, settings.languageRules, logger, context);
    }
    const hasLineBreaks = /[\r\n]/.test(str);
    if (logger && hasLineBreaks) {
        logger('Line breaks removed', str, str.replace(/[\r\n]+/g, ' '), context);
    }
    const withoutLineBreaks = str.replace(/[\r\n]+/g, ' ');
    const normalized = withoutLineBreaks.replace(/ {2,}/g, ' ').trim();
    if (logger && normalized !== withoutLineBreaks) {
        logger('Whitespace normalization', withoutLineBreaks, normalized, context);
    }
    return stripOuterQuotes(normalized, settings.stripOuterQuotes, logger, context);
}

// Convert XLSX data to TXT format
function xlsxToTxt(data) {
    // XLSX: rows = variables (Label1, Label2...), cols = languages
    // TXT: rows = languages, cols = variables + Date column

    resetLog('XLSX → TXT');

    // Transpose: now rows = languages, cols = variables
    const transposed = transpose(data);

    // Add Date column after first column and convert language codes
    const currentDate = getCurrentDate();
    const result = transposed.map((row, index) => {
        const newRow = row.map(cell => cell);
        let langCode = null;

        // Convert language code in first column (e.g., en_EN -> EN)
        if (index > 0 && newRow[0]) {
            langCode = localeToShort(newRow[0]);
            newRow[0] = langCode;
        }

        const normalizedRow = newRow.map((cell, cellIndex) => {
            const originalCol = index + 1;
            const originalRow = cellIndex + 1;
            const cellRef = `${columnToLetters(originalCol)}${originalRow}`;
            const context = { cellRef, lang: langCode };
            return normalizeTxtCell(cell, addLog, context, langCode);
        });

        // Apply NBSP rules to all values except the first column (key)
        if (index > 0 && normalizedRow[0]) {
            langCode = normalizedRow[0];
            for (let i = 1; i < normalizedRow.length; i++) {
                const originalCol = index + 1;
                const originalRow = i + 1;
                const cellRef = `${columnToLetters(originalCol)}${originalRow}`;
                const context = { cellRef, lang: langCode };
                normalizedRow[i] = applyNbspRules(normalizedRow[i], langCode, {
                    logger: (rule, before, after) => addLog(rule, before, after, context)
                });
            }
        }

        // Insert "Date" header or date value after first column
        if (index === 0) {
            normalizedRow.splice(1, 0, 'Date');
        } else {
            normalizedRow.splice(1, 0, currentDate);
        }

        return normalizedRow;
    });

    // Convert to tab-separated string
    const txt = result.map(row => row.join('\t')).join('\n');
    renderLog();
    return txt;
}

// Convert TXT data to XLSX format
function txtToXlsx(data) {
    // TXT: rows = languages, cols = variables + Date column
    // XLSX: rows = variables (without Date), cols = languages

    // Remove Date column (index 1) and convert language codes
    const settings = getTxtSettings();
    const withoutDate = data.map((row, index) => {
        const langCode = index > 0 && row[0] ? localeToShort(row[0]) : null;
        const isFrench = (langCode || "").toUpperCase() === 'FR';
        const newRow = row.map((cell, cellIndex) => {
            const cellRef = `${columnToLetters(cellIndex + 1)}${index + 1}`;
            const context = { cellRef, lang: langCode };
            let value = cell;
            if (isFrench) {
                value = applyFrenchNarrowSpaces(value, settings.languageRules, addLog, context);
            }
            value = stripOuterQuotes(value, settings.stripOuterQuotes, addLog, context);
            if (index > 0 && cellIndex > 0) {
                value = applyNbspRules(value, langCode, {
                    settings,
                    logger: (rule, before, after) => addLog(rule, before, after, context)
                });
            }
            return value;
        });
        newRow.splice(1, 1);
        // Convert language code in first column (e.g., EN -> en_EN)
        if (index > 0 && newRow[0]) {
            newRow[0] = shortToLocale(newRow[0]);
        }
        return newRow;
    });

    // Transpose: now rows = variables, cols = languages
    return transpose(withoutDate);
}

// Mac Roman encoding table: Unicode codepoint -> Mac Roman byte
const macRomanFromUnicode = {
    0x00C4: 0x80, 0x00C5: 0x81, 0x00C7: 0x82, 0x00C9: 0x83, 0x00D1: 0x84,
    0x00D6: 0x85, 0x00DC: 0x86, 0x00E1: 0x87, 0x00E0: 0x88, 0x00E2: 0x89,
    0x00E4: 0x8A, 0x00E3: 0x8B, 0x00E5: 0x8C, 0x00E7: 0x8D, 0x00E9: 0x8E,
    0x00E8: 0x8F, 0x00EA: 0x90, 0x00EB: 0x91, 0x00ED: 0x92, 0x00EC: 0x93,
    0x00EE: 0x94, 0x00EF: 0x95, 0x00F1: 0x96, 0x00F3: 0x97, 0x00F2: 0x98,
    0x00F4: 0x99, 0x00F6: 0x9A, 0x00F5: 0x9B, 0x00FA: 0x9C, 0x00F9: 0x9D,
    0x00FB: 0x9E, 0x00FC: 0x9F, 0x2020: 0xA0, 0x00B0: 0xA1, 0x00A2: 0xA2,
    0x00A3: 0xA3, 0x00A7: 0xA4, 0x2022: 0xA5, 0x00B6: 0xA6, 0x00DF: 0xA7,
    0x00AE: 0xA8, 0x00A9: 0xA9, 0x2122: 0xAA, 0x00B4: 0xAB, 0x00A8: 0xAC,
    0x2260: 0xAD, 0x00C6: 0xAE, 0x00D8: 0xAF, 0x221E: 0xB0, 0x00B1: 0xB1,
    0x2264: 0xB2, 0x2265: 0xB3, 0x00A5: 0xB4, 0x00B5: 0xB5, 0x2202: 0xB6,
    0x2211: 0xB7, 0x220F: 0xB8, 0x03C0: 0xB9, 0x222B: 0xBA, 0x00AA: 0xBB,
    0x00BA: 0xBC, 0x03A9: 0xBD, 0x00E6: 0xBE, 0x00F8: 0xBF, 0x00BF: 0xC0,
    0x00A1: 0xC1, 0x00AC: 0xC2, 0x221A: 0xC3, 0x0192: 0xC4, 0x2248: 0xC5,
    0x2206: 0xC6, 0x00AB: 0xC7, 0x00BB: 0xC8, 0x2026: 0xC9, 0x00A0: 0xCA,
    0x00C0: 0xCB, 0x00C3: 0xCC, 0x00D5: 0xCD, 0x0152: 0xCE, 0x0153: 0xCF,
    0x2013: 0xD0, 0x2014: 0xD1, 0x201C: 0xD2, 0x201D: 0xD3, 0x2018: 0xD4,
    0x2019: 0xD5, 0x00F7: 0xD6, 0x25CA: 0xD7, 0x00FF: 0xD8, 0x0178: 0xD9,
    0x2044: 0xDA, 0x20AC: 0xDB, 0x2039: 0xDC, 0x203A: 0xDD, 0xFB01: 0xDE,
    0xFB02: 0xDF, 0x2021: 0xE0, 0x00B7: 0xE1, 0x201A: 0xE2, 0x201E: 0xE3,
    0x2030: 0xE4, 0x00C2: 0xE5, 0x00CA: 0xE6, 0x00C1: 0xE7, 0x00CB: 0xE8,
    0x00C8: 0xE9, 0x00CD: 0xEA, 0x00CE: 0xEB, 0x00CF: 0xEC, 0x00CC: 0xED,
    0x00D3: 0xEE, 0x00D4: 0xEF, 0xF8FF: 0xF0, 0x00D2: 0xF1, 0x00DA: 0xF2,
    0x00DB: 0xF3, 0x00D9: 0xF4, 0x0131: 0xF5, 0x02C6: 0xF6, 0x02DC: 0xF7,
    0x00AF: 0xF8, 0x02D8: 0xF9, 0x02D9: 0xFA, 0x02DA: 0xFB, 0x00B8: 0xFC,
    0x02DD: 0xFD, 0x02DB: 0xFE, 0x02C7: 0xFF
};

// Windows-1252 encoding table: Unicode codepoint -> Windows-1252 byte
const win1252FromUnicode = {
    0x20AC: 0x80, 0x201A: 0x82, 0x0192: 0x83, 0x201E: 0x84, 0x2026: 0x85,
    0x2020: 0x86, 0x2021: 0x87, 0x02C6: 0x88, 0x2030: 0x89, 0x0160: 0x8A,
    0x2039: 0x8B, 0x0152: 0x8C, 0x017D: 0x8E, 0x2018: 0x91, 0x2019: 0x92,
    0x201C: 0x93, 0x201D: 0x94, 0x2022: 0x95, 0x2013: 0x96, 0x2014: 0x97,
    0x02DC: 0x98, 0x2122: 0x99, 0x0161: 0x9A, 0x203A: 0x9B, 0x0153: 0x9C,
    0x017E: 0x9E, 0x0178: 0x9F
};

// Encode string to specific encoding
function encodeString(str, encoding) {
    // For UTF-8, use native TextEncoder
    if (encoding === 'utf-8') {
        return new TextEncoder().encode(str);
    }

    const bytes = [];
    const encTable = encoding === 'macintosh' ? macRomanFromUnicode : win1252FromUnicode;

    for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);

        if (code < 0x80) {
            // ASCII - same in all encodings
            bytes.push(code);
        } else if (encTable[code] !== undefined) {
            // Use encoding table
            bytes.push(encTable[code]);
        } else if (code >= 0x00A0 && code <= 0x00FF) {
            // Latin-1 Supplement (mostly same in Windows-1252)
            bytes.push(code);
        } else {
            // Character not in encoding - use '?'
            bytes.push(0x3F);
        }
    }

    return new Uint8Array(bytes);
}

// Trigger TXT download with specified encoding
function downloadTXT(content, filename, encoding) {
    const bytes = encodeString(content, encoding);
    const blob = new Blob([bytes], { type: 'text/plain' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.textContent = `Download ${filename}`;
    return link;
}

// Trigger XLSX download
function downloadXLSX(data, filename) {
    const worksheet = XLSX.utils.aoa_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    const xlsxData = XLSX.write(workbook, { bookType: 'xlsx', type: 'array' });
    const blob = new Blob([xlsxData], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = filename;
    link.textContent = `Download ${filename}`;
    return link;
}

// Convert XLSX to TXT
async function convertXlsxToTxt() {
    try {
        xlsxError.textContent = '';
        txtDownload.innerHTML = '';

        const encoding = xlsxOutputEncoding.value;
        const data = await parseXLSX(xlsxFile);
        const txtContent = xlsxToTxt(data);

        const baseName = xlsxFile.name.replace(/\.xlsx$/i, '');
        const link = downloadTXT(txtContent, `${baseName}.txt`, encoding);
        txtDownload.appendChild(link);
    } catch (err) {
        xlsxError.textContent = 'Error: ' + err.message;
    }
}

// Convert TXT to XLSX
async function convertTxtToXlsx() {
    try {
        txtError.textContent = '';
        xlsxDownload.innerHTML = '';
        resetLog('TXT → XLSX');

        const encoding = txtInputEncoding.value;
        const data = await parseTXT(txtFile, encoding);
        const xlsxData = txtToXlsx(data);
        renderLog();

        const baseName = txtFile.name.replace(/\.txt$/i, '');
        const link = downloadXLSX(xlsxData, `${baseName}.xlsx`);
        xlsxDownload.appendChild(link);
    } catch (err) {
        txtError.textContent = 'Error: ' + err.message;
    }
}

