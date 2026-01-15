(() => {
    const NBSP = "\u00A0";
    const DEFAULT_SETTINGS = {
        numericFormatting: true,
        numberSeparators: true,
        percentFormatting: true,
        currencyFormatting: true,
        unitSpacing: true,
        unitCaseNormalization: true,
        globalRules: true,
        languageRules: true
    };

    function getSettings() {
        const current = window.nbspSettings || {};
        return { ...DEFAULT_SETTINGS, ...current };
    }
    const NUMBER_PATTERN = '(?:\\d[\\d.,\\u00A0 ]*\\d|\\d)';
    const UNIT_TOKENS = [
        'W', 'kW', 'MW', 'GW', 'hp', 'PS', 'CV',
        'V', 'kV', 'mV', 'A', 'mA', 'Hz', 'kHz', 'MHz', 'GHz', 'Ohm', 'Ω',
        'Wh', 'kWh', 'MWh', 'J', 'kJ', 'F', 'µF', 'μF', 'nF',
        'N', 'kN', 'Kgf', 'kgf', 'Pa', 'kPa', 'bar', 'Nm',
        'km', 'm', 'cm', 'mm', 'kg', 'g', 'mg', 'l', 'ml'
    ];

    function applyGlobalRules(text, enabled) {
        if (!enabled) return text;
        let result = text;

        // Initials and surname
        result = result.replace(/\b([A-ZÀ-ÖØ-Þ])\.\s+([A-ZÀ-ÖØ-Þ][A-Za-zÀ-ÖØ-öø-ÿ'-]*)/g, `$1.${NBSP}$2`);

        // Abbreviations and numerals
        result = result.replace(/\b(p\.|№|Vol\.)\s+(\d+)/gi, `$1${NBSP}$2`);

        // Dates: "Jan 15, 2026 year"
        result = result.replace(/\b(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2},)/g, `$1${NBSP}$2`);
        result = result.replace(/\b(\d{4})\s+(year)\b/gi, `$1${NBSP}$2`);

        return result;
    }

    function applyUnitSpacing(text, enabled) {
        if (!enabled) return text;

        const simpleUnits = `(?:${UNIT_TOKENS.join('|')})`;
        const compositeUnits = '(?:[A-Za-zΩµμ]+[0-9]*(?:[./][A-Za-zΩµμ0-9]+)+)';
        const spacedUnitsRegex = new RegExp(`(${NUMBER_PATTERN})\\s+(${simpleUnits}|${compositeUnits})\\b`, 'gi');
        const tightUnitsRegex = new RegExp(`(${NUMBER_PATTERN})(${simpleUnits}|${compositeUnits})\\b`, 'gi');

        let result = text.replace(spacedUnitsRegex, `$1${NBSP}$2`);
        result = result.replace(tightUnitsRegex, `$1${NBSP}$2`);
        return result;
    }

    function applyUnitCaseNormalization(text, enabled) {
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
            return `${num}${sep}${normalizedUnit}`;
        });
    }

    function applyLanguageRules(text, langCode, enabled) {
        if (!enabled) return text;
        const lang = (langCode || "").toUpperCase();
        let result = text;

        if (lang === "EN") {
            result = result.replace(/\b(a|an|the|I)\s+/gi, (match, word) => `${word}${NBSP}`);
        } else if (lang === "FR") {
            result = result.replace(/\b(à|y|en|le|la|les|un|une|de|du)\s+/gi, (match, word) => `${word}${NBSP}`);
            result = result.replace(/([^\s])\s*([;:!?])/g, `$1${NBSP}$2`);
            result = result.replace(/«\s*/g, `«${NBSP}`);
            result = result.replace(/\s*»/g, `${NBSP}»`);
        } else if (lang === "ES") {
            result = result.replace(/\b(a|e|i|o|u|y|la|el|un)\s+/gi, (match, word) => `${word}${NBSP}`);
        } else if (lang === "DE") {
            result = result.replace(/\b(der|die|das|ein|eine|Dr\.|Prof\.|Hr\.|Fr\.)\s+/gi, (match, word) => `${word}${NBSP}`);
        } else if (lang === "NL") {
            result = result.replace(/\b(de|het|een|in|op|te|ten|ter)\s+/gi, (match, word) => `${word}${NBSP}`);
        }

        return result;
    }

    function applyNumberSeparators(text, langCode, enabled) {
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

        function formatIntegerPart(intPart, separator) {
            return intPart.replace(/\B(?=(\d{3})+(?!\d))/g, separator);
        }

        const numberRegex = new RegExp(NUMBER_PATTERN, 'g');
        return text.replace(numberRegex, (match) => {
            const parts = parseNumberString(match);
            if (!parts) return match;
            const formattedInt = formatIntegerPart(parts.intPart, thousandsSep);
            if (!parts.fracPart) return formattedInt;
            return `${formattedInt}${decimalSep}${parts.fracPart}`;
        });
    }

    function formatPercent(text, langCode, enabled) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const numberRegex = new RegExp(`(${NUMBER_PATTERN})\\s*%`, 'g');
        if (lang === "EN") {
            return text.replace(numberRegex, `$1%`);
        }
        return text.replace(numberRegex, `$1${NBSP}%`);
    }

    function formatCurrency(text, langCode, enabled) {
        if (!enabled) return text;

        const lang = (langCode || "").toUpperCase();
        const numberCapture = NUMBER_PATTERN;
        const prefixRegex = new RegExp(`([€$£])\\s*(${numberCapture})`, 'g');
        const suffixRegex = new RegExp(`(${numberCapture})\\s*([€$£])`, 'g');

        if (lang === "EN") {
            let result = text.replace(suffixRegex, `$2$1`);
            result = result.replace(prefixRegex, `$1$2`);
            return result;
        }

        if (lang === "NL") {
            let result = text.replace(suffixRegex, `$2${NBSP}$1`);
            result = result.replace(prefixRegex, `$1${NBSP}$2`);
            return result;
        }

        if (lang === "FR" || lang === "DE" || lang === "ES") {
            let result = text.replace(prefixRegex, `$2${NBSP}$1`);
            result = result.replace(suffixRegex, `$1${NBSP}$2`);
            return result;
        }

        return text;
    }

    function applyNumericFormatting(text, langCode, enabled) {
        if (!enabled) return text;
        const settings = getSettings();
        let result = applyNumberSeparators(text, langCode, settings.numberSeparators);
        result = formatPercent(result, langCode, settings.percentFormatting);
        result = formatCurrency(result, langCode, settings.currencyFormatting);
        return result;
    }

    window.applyNbspRules = function applyNbspRules(value, langCode) {
        if (value === null || value === undefined) return '';
        const text = String(value);
        if (!text) return text;

        const settings = getSettings();
        const withNumeric = applyNumericFormatting(text, langCode, settings.numericFormatting);
        const withUnits = applyUnitSpacing(withNumeric, settings.unitSpacing);
        const withUnitCase = applyUnitCaseNormalization(withUnits, settings.unitCaseNormalization);
        const withGlobal = applyGlobalRules(withUnitCase, settings.globalRules);
        return applyLanguageRules(withGlobal, langCode, settings.languageRules);
    };
})();

