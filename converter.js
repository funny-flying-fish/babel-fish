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
        result = result.replace(/\b([A-ZÀ-ÖØ-ÞĄĆĘŁŃŚŹŻ])\.\s+([A-ZÀ-ÖØ-ÞĄĆĘŁŃŚŹŻ][A-Za-zÀ-ÖØ-öø-ÿĄĆĘŁŃŚŹŻąćęłńśźż'-]*)/g, (match, initial, surname) => {
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

        // \b is broken for Unicode: JS treats only [A-Za-z0-9_] as word chars,
        // so \b fires inside words like "Contrôle", "systemów", "autonomía".
        // Use negative lookbehind for any Latin letter (including diacritics) instead.
        const LB = '(?<![A-Za-z\\xC0-\\xD6\\xD8-\\xF6\\xF8-\\xFF\\u0100-\\u024F])';

        if (lang === "EN") {
            result = result.replace(new RegExp(LB + '(a|an|the|I)\\s+', 'gi'), (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'EN NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "FR") {
            result = result.replace(new RegExp(LB + '(à|y|en|le|la|les|un|une|de|du)\\s+', 'gi'), (match, word) => {
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
            result = result.replace(new RegExp(LB + '(a|e|i|o|u|y|la|el|un)\\s+', 'gi'), (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'ES NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "DE") {
            result = result.replace(new RegExp(LB + '(der|die|das|ein|eine|Dr\\.|Prof\\.|Hr\\.|Fr\\.)\\s+', 'gi'), (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'DE NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "NL") {
            result = result.replace(new RegExp(LB + '(de|het|een|in|op|te|ten|ter)\\s+', 'gi'), (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'NL NBSP after short words', match, updated);
                return updated;
            });
        } else if (lang === "PL") {
            result = result.replace(new RegExp(LB + '(a|e|i|o|u|w|z)\\s+', 'gi'), (match, word) => {
                const updated = `${word}${NBSP}`;
                logChange(logger, 'PL NBSP after short words', match, updated);
                return updated;
            });
        }

        return result;
    }

    function applyNumberSeparators(text, langCode, enabled, logger) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const thousandsSep = (lang === "FR" || lang === "PL") ? NBSP : (lang === "EN" ? "," : ".");
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

                const beforeDigits = before.replace(/[.,\s]/g, '');
                if (!(sepCount === 1 && digitsAfter === 3) || beforeDigits === '0') {
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

        if (lang === "FR" || lang === "DE" || lang === "ES" || lang === "PL") {
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

// File references (arrays for multi-file support)
let xlsxFiles = [];
let txtFiles = [];

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
const xlsxDropzone = document.getElementById('xlsxDropzone');
const txtDropzone = document.getElementById('txtDropzone');
const xlsxFileList = document.getElementById('xlsxFileList');
const txtFileList = document.getElementById('txtFileList');
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
const ruleDuplicateCheck = document.getElementById('ruleDuplicateCheck');
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
let currentLogFilename = '';

function addLog(rule, before, after, context) {
    logEntries.push({ rule, before, after, context, filename: currentLogFilename });
}

function addLogNotice(message) {
    if (!message) return;
    logEntries.push({ rule: 'Notice', before: message, after: '', context: null, filename: currentLogFilename });
}

function addLogError(message) {
    if (!message) return;
    logEntries.push({ rule: 'Error', before: message, after: '', context: null, filename: currentLogFilename });
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
    const changeCount = logEntries.filter(e => e.rule !== 'Notice').length;
    logCount.textContent = `${changeCount} changes`;
    logEntries.forEach((entry) => {
        const item = document.createElement('li');

        // Build prefix: filename + cell ref
        const prefixParts = [];
        if (entry.filename) {
            prefixParts.push(`<span class="log-file-name">${escapeHtml(entry.filename)}</span>`);
        }
        if (entry.context && entry.context.cellRef) {
            prefixParts.push(formatCellRefHtml(entry.context.cellRef));
        }
        const prefixHtml = prefixParts.length ? `${prefixParts.join(' ')} ` : '';

        if (entry.rule === 'Notice' || entry.rule === 'Error') {
            const ruleClass = entry.rule === 'Error' ? 'log-rule log-error' : 'log-rule';
            item.innerHTML = `${prefixHtml}<span class="${ruleClass}">${escapeHtml(entry.rule)}</span>: ` +
                `<span class="${entry.rule === 'Error' ? 'log-error' : ''}">${escapeHtml(entry.before)}</span>`;
        } else if (entry.rule === 'Whitespace normalization' || entry.rule === 'Line breaks removed') {
            const valueHtml = entry.rule === 'Line breaks removed'
                ? formatLineBreakValue(entry.before)
                : formatWhitespaceValue(entry.before);
            item.innerHTML = `${prefixHtml}<span class="log-rule">${escapeHtml(entry.rule)}</span>: ` +
                `<span class="log-line">"${valueHtml}"</span>`;
        } else {
            const beforeHtml = escapeHtml(entry.before);
            const afterHtml = escapeHtml(entry.after);
            item.innerHTML = `${prefixHtml}<span class="log-rule">${escapeHtml(entry.rule)}</span>: ` +
                `<span class="log-before">"${beforeHtml}"</span> → ` +
                `<span class="log-after">"${afterHtml}"</span>`;
        }
        logList.appendChild(item);
    });
}

function resetLog(title) {
    logEntries = [];
    currentLogFilename = '';
    logTitle.textContent = title;
    renderLog();
}

// ── Dropzone helpers ─────────────────────────────────────────────
function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function renderFileList(files, listEl, dropzone) {
    listEl.innerHTML = '';
    if (files.length === 0) {
        dropzone.classList.remove('has-files');
        return;
    }
    dropzone.classList.add('has-files');
    files.forEach((file, index) => {
        const li = document.createElement('li');
        li.className = 'dropzone-file-item';
        li.innerHTML = `
            <span class="file-name">${escapeHtml(file.name)}</span>
            <span class="file-size">${formatFileSize(file.size)}</span>
            <button type="button" class="file-remove" data-index="${index}" title="Remove">✕</button>
        `;
        listEl.appendChild(li);
    });
}

function addFilesToList(newFiles, currentFiles, accept) {
    const ext = accept.replace('.', '').toLowerCase();
    const filtered = Array.from(newFiles).filter(f => f.name.toLowerCase().endsWith('.' + ext));
    // Avoid duplicates by name+size
    const existing = new Set(currentFiles.map(f => f.name + '|' + f.size));
    for (const f of filtered) {
        if (!existing.has(f.name + '|' + f.size)) {
            currentFiles.push(f);
        }
    }
    return currentFiles;
}

function setupDropzone(dropzone, input, getFiles, setFiles, listEl, btn, downloadArea, errorEl, accept) {
    // Click to browse
    dropzone.addEventListener('click', (e) => {
        if (e.target.closest('.file-remove')) return;
        input.click();
    });

    // Drag events
    dropzone.addEventListener('dragenter', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', (e) => {
        if (!dropzone.contains(e.relatedTarget)) dropzone.classList.remove('dragover');
    });
    dropzone.addEventListener('drop', (e) => {
        e.preventDefault();
        dropzone.classList.remove('dragover');
        const files = addFilesToList(e.dataTransfer.files, getFiles(), accept);
        setFiles(files);
        renderFileList(files, listEl, dropzone);
        btn.disabled = files.length === 0;
        downloadArea.innerHTML = '';
        errorEl.textContent = '';
    });

    // File input change
    input.addEventListener('change', () => {
        const files = addFilesToList(input.files, getFiles(), accept);
        setFiles(files);
        renderFileList(files, listEl, dropzone);
        btn.disabled = files.length === 0;
        downloadArea.innerHTML = '';
        errorEl.textContent = '';
        input.value = '';
    });

    // Remove file button (delegated)
    listEl.addEventListener('click', (e) => {
        const removeBtn = e.target.closest('.file-remove');
        if (!removeBtn) return;
        e.stopPropagation();
        const idx = parseInt(removeBtn.dataset.index, 10);
        const files = getFiles();
        files.splice(idx, 1);
        setFiles(files);
        renderFileList(files, listEl, dropzone);
        btn.disabled = files.length === 0;
    });
}

setupDropzone(
    xlsxDropzone, xlsxInput,
    () => xlsxFiles, (f) => { xlsxFiles = f; },
    xlsxFileList, xlsxToTxtBtn, txtDownload, xlsxError, '.xlsx'
);

setupDropzone(
    txtDropzone, txtInput,
    () => txtFiles, (f) => { txtFiles = f; },
    txtFileList, txtToXlsxBtn, xlsxDownload, txtError, '.txt'
);

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

function extractCodeFromFilename(filename) {
    if (!filename) return { value: 'not-set', hasCode: false };
    const squareMatch = filename.match(/\[(.*?)\]/);
    if (squareMatch) {
        const trimmed = squareMatch[1].trim();
        if (trimmed) return { value: trimmed, hasCode: true };
    }
    const roundMatch = filename.match(/\((.*?)\)/);
    if (roundMatch) {
        const trimmed = roundMatch[1].trim();
        if (trimmed) return { value: trimmed, hasCode: true };
    }
    return { value: 'not-set', hasCode: false };
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

function detectEncoding(buffer) {
    const bytes = new Uint8Array(buffer);
    if (bytes.length >= 2 && bytes[0] === 0xFF && bytes[1] === 0xFE) return 'utf-16le';
    if (bytes.length >= 3 && bytes[0] === 0xEF && bytes[1] === 0xBB && bytes[2] === 0xBF) return 'utf-8';
    // Detect BOM-less UTF-16LE: check if odd-position bytes are mostly 0x00
    // (typical for ASCII/Latin text encoded as UTF-16LE without a BOM)
    if (bytes.length >= 20) {
        let nullCount = 0;
        const checkLen = Math.min(bytes.length, 100);
        for (let i = 1; i < checkLen; i += 2) {
            if (bytes[i] === 0x00) nullCount++;
        }
        if (nullCount > (checkLen / 2) * 0.8) return 'utf-16le';
    }
    return 'macintosh';
}

// Parse TXT file with auto-detected encoding and return 2D array
function parseTXT(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const buffer = e.target.result;
                const encoding = detectEncoding(buffer);
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

function normalizeNonBreakingHyphen(text, enabled, logger, context) {
    if (text === null || text === undefined) return '';
    const value = String(text);
    if (!enabled) return value;
    const normalized = value.replace(/\u2011/g, '-');
    if (logger && normalized !== value) {
        logger('Non-breaking hyphen normalized', value, normalized, context);
    }
    return normalized;
}

function normalizeTxtCell(value, logger, context, langCode, options = {}) {
    if (value === null || value === undefined) return '';
    let str = String(value);
    const settings = getSettings();
    const lang = (langCode || "").toUpperCase();
    str = normalizeNonBreakingHyphen(str, options.replaceNonBreakingHyphen, logger, context);
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

// Language code detection for key row/column
const KNOWN_LANG_CODES = new Set([
    'EN', 'FR', 'DE', 'ES', 'NL', 'IT', 'PT', 'RU', 'PL', 'CS',
    'SK', 'HU', 'RO', 'BG', 'HR', 'SL', 'SR', 'BS', 'MK', 'SQ',
    'EL', 'TR', 'DA', 'NO', 'SV', 'FI', 'ET', 'LV', 'LT', 'GA',
    'CY', 'MT', 'EU', 'CA', 'GL', 'JA', 'ZH', 'KO', 'AR', 'HE',
    'HI', 'TH', 'VI', 'ID', 'MS', 'TL', 'SW', 'UK', 'BE'
]);

function isLanguageCode(str) {
    if (!str) return false;
    const trimmed = String(str).trim();
    // Locale format: en_EN, fr_FR, etc. — very reliable
    if (/^[a-z]{2}_[A-Z]{2}$/.test(trimmed)) return true;
    // Short format: EN, FR, en, fr — validate against known codes
    const upper = trimmed.toUpperCase();
    if (/^[A-Za-z]{2}$/.test(trimmed) && KNOWN_LANG_CODES.has(upper)) return true;
    return false;
}

function hasLanguageKeyRow(data) {
    if (!data || data.length === 0) return false;
    const firstRow = data[0];
    // Check cells from index 1 onwards (index 0 is the key/label column header)
    for (let i = 1; i < firstRow.length; i++) {
        if (isLanguageCode(firstRow[i])) return true;
    }
    return false;
}

function hasLanguageKeyColumn(data, colIndex) {
    if (!data || data.length <= 1) return false;
    // Check data rows (skip header row at index 0)
    for (let i = 1; i < data.length; i++) {
        if (data[i] && isLanguageCode(data[i][colIndex])) return true;
    }
    return false;
}

// Convert XLSX data to TXT format
function xlsxToTxt(data, options = {}) {
    // XLSX: rows = variables (Label1, Label2...), cols = languages
    // TXT: rows = languages, cols = variables + Date column
    const replaceNonBreakingHyphen = Boolean(options.replaceNonBreakingHyphen);
    let foundNonBreakingHyphen = false;
    if (options.codeHasValue === false) {
        addLogNotice('No value in [] in filename. Using code = not-set.');
    }

    // Check if XLSX has a language key row (row 0 should have language codes)
    if (!hasLanguageKeyRow(data)) {
        const maxCols = data.length > 0 ? data.reduce((max, row) => Math.max(max, row.length), 0) : 2;
        const keyRow = [''];
        for (let i = 1; i < maxCols; i++) {
            keyRow.push('en_EN');
        }
        data.unshift(keyRow);
        addLogNotice('No language key row found in XLSX. Added default language "EN".');
    }

    // Transpose: now rows = languages, cols = variables
    const transposed = transpose(data);

    // Add Date column after first column and convert language codes
    const currentDate = getCurrentDate();
    const codeValue = options.codeValue || 'not-set';
    const result = transposed.map((row, index) => {
        const newRow = row.map(cell => cell);
        let langCode = null;

        // Convert language code in first column (e.g., en_EN -> EN)
        if (index > 0 && newRow[0]) {
            langCode = localeToShort(newRow[0]);
            newRow[0] = String(langCode).toUpperCase();
        }

        const normalizedRow = newRow.map((cell, cellIndex) => {
            const originalCol = index + 1;
            const originalRow = cellIndex + 1;
            const cellRef = `${columnToLetters(originalCol)}${originalRow}`;
            const context = { cellRef, lang: langCode };
            if (cell !== null && cell !== undefined && String(cell).includes('\u2011')) {
                foundNonBreakingHyphen = true;
            }
            return normalizeTxtCell(cell, addLog, context, langCode, {
                replaceNonBreakingHyphen
            });
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

        // Insert "code" column as the first column
        if (index === 0) {
            normalizedRow.unshift('code');
        } else {
            normalizedRow.unshift(codeValue);
        }

        return normalizedRow;
    });

    // Convert to tab-separated string
    const txt = result.map(row => row.join('\t')).join('\n');
    return { txt, foundNonBreakingHyphen };
}

// Convert TXT data to XLSX format
function txtToXlsx(data) {
    // TXT: rows = languages, cols = variables + Date column
    // XLSX: rows = variables (without Date), cols = languages

    const settings = getTxtSettings();
    const hasCodeColumn = data.length > 0 && String(data[0][0] || '').trim().toLowerCase() === 'code';

    // Determine language column index in original data
    const langColIndex = hasCodeColumn ? 1 : 0;

    // Check if TXT has a language key column
    if (!hasLanguageKeyColumn(data, langColIndex)) {
        for (let i = 0; i < data.length; i++) {
            if (i === 0) {
                data[i].splice(langColIndex, 0, 'Key');
            } else {
                data[i].splice(langColIndex, 0, 'EN');
            }
        }
        addLogNotice('No language key column found in TXT. Added default language "EN".');
    }

    // Set language column header to "Key" for XLSX output (becomes cell A1)
    if (data.length > 0 && data[0]) {
        data[0][langColIndex] = 'Key';
    }

    // Detect Date column: check effective header at index 1
    const effectiveHeader = hasCodeColumn ? data[0].slice(1) : [...data[0]];
    const hasDateColumn = effectiveHeader.length > 1 &&
        String(effectiveHeader[1] || '').trim().toLowerCase() === 'date';

    // Check for duplicate language keys in data cells (when checkbox is checked)
    let duplicateWarning = null;
    if (ruleDuplicateCheck && ruleDuplicateCheck.checked) {
        // Collect language codes from the language column
        const langCodesInColumn = new Set();
        for (let i = 1; i < data.length; i++) {
            const code = String(data[i][langColIndex] || '').trim().toUpperCase();
            if (code) langCodesInColumn.add(code);
        }

        // 1) Check for duplicate language rows (same code in multiple rows)
        const langCounts = {};
        for (let i = 1; i < data.length; i++) {
            const code = String(data[i][langColIndex] || '').trim().toUpperCase();
            if (!code) continue;
            if (!langCounts[code]) langCounts[code] = [];
            langCounts[code].push(i);
        }

        let keyOldCount = 0;
        for (const [code, indices] of Object.entries(langCounts)) {
            if (indices.length > 1) {
                addLogError(`Duplicate language row "${code}" found in rows ${indices.map(i => i + 1).join(', ')}. Renaming duplicates to "key-old".`);
                duplicateWarning = duplicateWarning || `Duplicate language key "${code}" found. Duplicates renamed to "key-old".`;
                for (let j = 1; j < indices.length; j++) {
                    keyOldCount++;
                    const newName = keyOldCount === 1 ? 'key-old' : `key-old-${keyOldCount}`;
                    data[indices[j]][langColIndex] = newName;
                }
            }
        }

        // 2) Check data cells for values matching a language code (e.g. "EN" or "EN ")
        // Data columns start after language column (+ Date column if present)
        const dataStartCol = langColIndex + (hasDateColumn ? 2 : 1);
        const flaggedColumns = new Set();
        for (let i = 1; i < data.length; i++) {
            for (let j = dataStartCol; j < (data[i] || []).length; j++) {
                const cellValue = String(data[i][j] || '').trim();
                if (cellValue && langCodesInColumn.has(cellValue.toUpperCase()) && !flaggedColumns.has(j)) {
                    keyOldCount++;
                    const newName = keyOldCount === 1 ? 'key-old' : `key-old-${keyOldCount}`;
                    const oldName = String(data[0][j] || '').trim() || `column ${j + 1}`;
                    data[0][j] = newName;
                    flaggedColumns.add(j);
                    addLogError(`Data cell contains language code "${cellValue}" in column "${oldName}" (row ${i + 1}). Column header renamed to "${newName}".`);
                    duplicateWarning = duplicateWarning || `Duplicate language key "${cellValue}" found in data column "${oldName}". Renamed to "${newName}".`;
                }
            }
        }
    }

    // Process rows: remove code column, remove Date column (only if present), convert language codes
    const processed = data.map((row, index) => {
        const effectiveRow = hasCodeColumn ? row.slice(1) : row;
        const langCode = index > 0 && effectiveRow[0] ? localeToShort(effectiveRow[0]) : null;
        const isFrench = (langCode || "").toUpperCase() === 'FR';
        const newRow = effectiveRow.map((cell, cellIndex) => {
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
        // Only remove Date column if it was detected
        if (hasDateColumn) {
            newRow.splice(1, 1);
        }
        // Convert language code in first column (e.g., EN -> en_EN)
        if (index > 0 && newRow[0]) {
            newRow[0] = shortToLocale(newRow[0]);
        }
        return newRow;
    });

    // Transpose: now rows = variables, cols = languages
    return { xlsxData: transpose(processed), duplicateWarning };
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

// Encode string to UTF-16 LE (each JS char → 2 bytes, little-endian)
function encodeUTF16LE(str) {
    const bytes = new Uint8Array(str.length * 2);
    for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);
        bytes[i * 2] = code & 0xFF;
        bytes[i * 2 + 1] = (code >> 8) & 0xFF;
    }
    return bytes;
}

// Encode string to specific encoding
function encodeString(str, encoding) {
    if (encoding === 'utf-8') {
        return new TextEncoder().encode(str);
    }

    if (encoding === 'utf-16le') {
        return encodeUTF16LE(str);
    }

    const bytes = [];
    const encTable = encoding === 'macintosh' ? macRomanFromUnicode : win1252FromUnicode;

    for (let i = 0; i < str.length; i++) {
        const code = str.charCodeAt(i);

        if (code < 0x80) {
            bytes.push(code);
        } else if (encTable[code] !== undefined) {
            bytes.push(encTable[code]);
        } else if (code >= 0x00A0 && code <= 0x00FF) {
            bytes.push(code);
        } else {
            bytes.push(0x3F);
        }
    }

    return new Uint8Array(bytes);
}

// Trigger TXT download with specified encoding
function downloadTXT(content, filename, encoding) {
    const bytes = encodeString(content, encoding);
    const parts = encoding === 'utf-16le'
        ? [new Uint8Array([0xFF, 0xFE]), bytes]
        : [bytes];
    const blob = new Blob(parts, { type: 'text/plain' });
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

// Convert XLSX to TXT (multi-file)
async function convertXlsxToTxt() {
    xlsxError.textContent = '';
    txtDownload.innerHTML = '';
    resetLog('XLSX → TXT log');

    if (xlsxFiles.length === 0) return;

    const encoding = xlsxOutputEncoding.value;
    let errorMessages = [];

    for (let fi = 0; fi < xlsxFiles.length; fi++) {
        const file = xlsxFiles[fi];
        currentLogFilename = file.name;
        try {
            const data = await parseXLSX(file);
            const { value: codeValue, hasCode } = extractCodeFromFilename(file.name);
            const { txt, foundNonBreakingHyphen } = xlsxToTxt(data, {
                replaceNonBreakingHyphen: encoding !== 'utf-16le',
                codeValue,
                codeHasValue: hasCode
            });

            if (foundNonBreakingHyphen && encoding !== 'utf-16le') {
                addLogNotice('Non-breaking hyphen found. Use Unicode output to preserve it. Otherwise the character will be replaced with a regular hyphen, which is safe.');
            }

            const encodingLabels = { 'utf-16le': 'Unicode', 'macintosh': 'MacRoman', 'windows-1252': 'Win1252' };
            const encodingLabel = encodingLabels[encoding] || encoding;
            const baseName = file.name.replace(/\.xlsx$/i, '');
            const link = downloadTXT(txt, `${baseName}-${encodingLabel}.txt`, encoding);
            txtDownload.appendChild(link);
        } catch (err) {
            errorMessages.push(`${file.name}: ${err.message}`);
            addLogError(`${file.name}: ${err.message}`);
        }
    }

    renderLog();
    if (errorMessages.length > 0) {
        xlsxError.textContent = 'Errors: ' + errorMessages.join('; ');
    }
}

// Convert TXT to XLSX (multi-file)
async function convertTxtToXlsx() {
    txtError.textContent = '';
    xlsxDownload.innerHTML = '';
    resetLog('TXT → XLSX log');

    if (txtFiles.length === 0) return;

    let errorMessages = [];
    let warnings = [];

    for (let fi = 0; fi < txtFiles.length; fi++) {
        const file = txtFiles[fi];
        currentLogFilename = file.name;
        try {
            const data = await parseTXT(file);
            const result = txtToXlsx(data);

            if (result.duplicateWarning) {
                warnings.push(`${file.name}: ${result.duplicateWarning}`);
            }

            const baseName = file.name.replace(/\.txt$/i, '');
            const link = downloadXLSX(result.xlsxData, `${baseName}.xlsx`);
            xlsxDownload.appendChild(link);
        } catch (err) {
            errorMessages.push(`${file.name}: ${err.message}`);
            addLogError(`${file.name}: ${err.message}`);
        }
    }

    renderLog();
    const allMessages = [...warnings.map(w => 'Warning: ' + w), ...errorMessages.map(e => 'Error: ' + e)];
    if (allMessages.length > 0) {
        txtError.innerHTML = allMessages.map(m => escapeHtml(m)).join('<br>');
    }
}

// ── Find difference ──────────────────────────────────────────────
let diffFileA = null;
let diffFileB = null;

const diffFileAInput = document.getElementById('diffFileA');
const diffFileBInput = document.getElementById('diffFileB');
const diffCompareBtn = document.getElementById('diffCompareBtn');
const diffError = document.getElementById('diffError');

function updateDiffButton() {
    diffCompareBtn.disabled = !(diffFileA && diffFileB);
}

diffFileAInput.addEventListener('change', (e) => {
    diffFileA = e.target.files[0] || null;
    updateDiffButton();
    diffError.textContent = '';
});

diffFileBInput.addEventListener('change', (e) => {
    diffFileB = e.target.files[0] || null;
    updateDiffButton();
    diffError.textContent = '';
});

diffCompareBtn.addEventListener('click', compareDiffFiles);

async function compareDiffFiles() {
    try {
        diffError.textContent = '';
        resetLog('Difference log');

        const dataA = await parseXLSX(diffFileA);
        const dataB = await parseXLSX(diffFileB);

        // Build maps: key (col A) -> english text (col B), skip header row
        const mapA = new Map();
        const mapB = new Map();
        const duplicateKeysA = [];
        const duplicateKeysB = [];

        for (let i = 1; i < dataA.length; i++) {
            const row = dataA[i];
            const key = String(row[0] || '').trim();
            if (!key) continue;
            if (mapA.has(key)) {
                duplicateKeysA.push({ key, row: i + 1 });
            }
            mapA.set(key, String(row[1] || '').trim());
        }

        for (let i = 1; i < dataB.length; i++) {
            const row = dataB[i];
            const key = String(row[0] || '').trim();
            if (!key) continue;
            if (mapB.has(key)) {
                duplicateKeysB.push({ key, row: i + 1 });
            }
            mapB.set(key, String(row[1] || '').trim());
        }

        const nameA = diffFileA.name;
        const nameB = diffFileB.name;

        // Report duplicate keys
        for (const dup of duplicateKeysA) {
            addLogError(`Duplicate key "${dup.key}" in ${nameA} (row ${dup.row})`);
        }
        for (const dup of duplicateKeysB) {
            addLogError(`Duplicate key "${dup.key}" in ${nameB} (row ${dup.row})`);
        }

        // 1) Key count difference
        if (mapA.size !== mapB.size) {
            addLog(
                'Key count mismatch',
                `${nameA}: ${mapA.size} keys`,
                `${nameB}: ${mapB.size} keys`
            );
        } else {
            addLogNotice(`Both files have ${mapA.size} keys.`);
        }

        // 2) Keys only in A
        const onlyInA = [];
        for (const key of mapA.keys()) {
            if (!mapB.has(key)) onlyInA.push(key);
        }
        if (onlyInA.length > 0) {
            addLogNotice(`${onlyInA.length} key(s) only in ${nameA}:`);
            for (const key of onlyInA) {
                addLog('Only in ' + nameA, key, '—');
            }
        }

        // 3) Keys only in B
        const onlyInB = [];
        for (const key of mapB.keys()) {
            if (!mapA.has(key)) onlyInB.push(key);
        }
        if (onlyInB.length > 0) {
            addLogNotice(`${onlyInB.length} key(s) only in ${nameB}:`);
            for (const key of onlyInB) {
                addLog('Only in ' + nameB, key, '—');
            }
        }

        // 4) Keys present in both but with different English text
        const textDiffs = [];
        for (const [key, textA] of mapA) {
            if (!mapB.has(key)) continue;
            const textB = mapB.get(key);
            if (textA !== textB) {
                textDiffs.push({ key, textA, textB });
            }
        }
        if (textDiffs.length > 0) {
            addLogNotice(`${textDiffs.length} key(s) with different English text:`);
            for (const diff of textDiffs) {
                addLog(diff.key, diff.textA, diff.textB);
            }
        }

        // Summary if no differences
        if (onlyInA.length === 0 && onlyInB.length === 0 && textDiffs.length === 0 &&
            duplicateKeysA.length === 0 && duplicateKeysB.length === 0 &&
            mapA.size === mapB.size) {
            addLogNotice('Files are identical — no differences found.');
        }

        renderLog();
    } catch (err) {
        diffError.textContent = 'Error: ' + err.message;
    }
}

