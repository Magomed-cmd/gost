"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.tagElement = tagElement;
exports.lintChildren = lintChildren;
exports.printLintResults = printLintResults;
const docx_1 = require("docx");
// ── Тегирование ───────────────────────────────────────────────────────────────
/**
 * Помечает элемент docx типом — используется внутри фабричных функций.
 * Возвращает тот же объект, не создаёт копию.
 */
function tagElement(el, type) {
    el.__gostType = type;
    return el;
}
function getType(el) {
    if (el instanceof docx_1.Table)
        return "table";
    return el.__gostType ?? "unknown";
}
/** Правило: pageBreak в самом начале — пустая страница перед содержимым */
const rulePageBreakAtStart = (types, i, issues) => {
    if (i !== 0)
        return;
    if (types[0] === "pageBreak") {
        issues.push({
            level: "warn", index: 0, code: "empty-page-start",
            message: "pageBreak на позиции 0 — пустая страница в начале документа.",
        });
    }
};
/** Правило: pageBreak в самом конце — пустая страница в конце документа */
const rulePageBreakAtEnd = (types, i, issues) => {
    const last = types.length - 1;
    if (i !== last)
        return;
    if (types[last] === "pageBreak") {
        issues.push({
            level: "warn", index: last, code: "empty-page-end",
            message: `pageBreak на позиции ${last} — пустая страница в конце документа.`,
        });
    }
};
/** Правило: два pageBreak подряд — заведомо пустая страница */
const ruleConsecutivePageBreaks = (types, i, issues) => {
    if (types[i] === "pageBreak" && types[i + 1] === "pageBreak") {
        issues.push({
            level: "warn", index: i, code: "consecutive-page-breaks",
            message: `Два pageBreak подряд на позиции ${i}–${i + 1} — пустая страница.`,
        });
    }
};
/** Правило: pageBreak → blank(и) → pageBreak — пустая страница с пробелами */
const rulePageBreakBlankPageBreak = (types, i, issues) => {
    if (types[i] !== "pageBreak")
        return;
    let j = i + 1;
    while (j < types.length && types[j] === "blank")
        j++;
    const blankCount = j - i - 1;
    if (blankCount > 0 && j < types.length && types[j] === "pageBreak") {
        issues.push({
            level: "warn", index: i, code: "page-break-blank-page-break",
            message: `pageBreak → ${blankCount} blank → pageBreak на позиции ${i}–${j} — пустая страница.`,
        });
    }
};
/** Правило: заголовок прямо перед pageBreak — заголовок повиснет внизу предыдущей страницы */
const ruleHeadingBeforePageBreak = (types, i, issues) => {
    const t = types[i];
    if ((t === "h1" || t === "h2" || t === "h3") && types[i + 1] === "pageBreak") {
        issues.push({
            level: "warn", index: i, code: "heading-before-page-break",
            message: `${t.toUpperCase()} на позиции ${i} сразу перед pageBreak — заголовок окажется один внизу страницы.`,
        });
    }
};
/** Правило: blank прямо перед pageBreak — лишний элемент */
const ruleBlankBeforePageBreak = (types, i, issues) => {
    if (types[i] === "blank" && types[i + 1] === "pageBreak") {
        issues.push({
            level: "info", index: i, code: "blank-before-page-break",
            message: `blank() на позиции ${i} перед pageBreak — лишний, не влияет на вёрстку.`,
        });
    }
};
/** Правило: 3+ blank подряд — может неожиданно сдвинуть контент */
const ruleTooManyBlanks = (types, i, issues) => {
    if (types[i] !== "blank" || types[i - 1] === "blank")
        return;
    let count = 0;
    let j = i;
    while (j < types.length && types[j] === "blank") {
        count++;
        j++;
    }
    if (count >= 3) {
        issues.push({
            level: "info", index: i, code: "too-many-blanks",
            message: `${count} blank() подряд на позиции ${i}–${j - 1} — может неожиданно перенести контент на следующую страницу.`,
        });
    }
};
/** Правило: image без caption следом — подпись обязательна по ГОСТ */
const ruleImageWithoutCaption = (types, i, issues) => {
    if (types[i] !== "image")
        return;
    const next = types[i + 1];
    if (next !== "caption") {
        issues.push({
            level: "warn", index: i, code: "image-without-caption",
            message: `image на позиции ${i} не имеет caption следом (ГОСТ требует подпись под каждым рисунком).`,
        });
    }
};
/** Правило: placeholder в документе — скорее всего забыли заменить на реальный рисунок */
const rulePlaceholderLeft = (types, i, issues) => {
    if (types[i] === "placeholder") {
        issues.push({
            level: "warn", index: i, code: "placeholder-left",
            message: `placeholder на позиции ${i} — не забудь заменить на реальное изображение.`,
        });
    }
};
const RULES = [
    rulePageBreakAtStart,
    rulePageBreakAtEnd,
    ruleConsecutivePageBreaks,
    rulePageBreakBlankPageBreak,
    ruleHeadingBeforePageBreak,
    ruleBlankBeforePageBreak,
    ruleTooManyBlanks,
    ruleImageWithoutCaption,
    rulePlaceholderLeft,
];
// ── Основная функция ──────────────────────────────────────────────────────────
/**
 * Анализирует массив children перед makeDocument и возвращает список проблем.
 * Не бросает исключений — только собирает предупреждения.
 */
function lintChildren(children) {
    const issues = [];
    const types = children.map(getType);
    for (let i = 0; i < types.length; i++) {
        for (const rule of RULES) {
            rule(types, i, issues);
        }
    }
    // Дедупликация: убираем дубли с одинаковым code + index
    const seen = new Set();
    return issues.filter(issue => {
        const key = `${issue.code}:${issue.index}`;
        if (seen.has(key))
            return false;
        seen.add(key);
        return true;
    });
}
/**
 * Выводит результаты lint в консоль в читаемом виде.
 * Вызывать перед saveDocument.
 */
function printLintResults(issues) {
    if (issues.length === 0) {
        console.log("✓ lintChildren: проблем не найдено.");
        return;
    }
    const warns = issues.filter(i => i.level === "warn");
    const infos = issues.filter(i => i.level === "info");
    for (const issue of issues) {
        const icon = issue.level === "warn" ? "⚠" : "ℹ";
        const label = issue.level === "warn" ? "WARN" : "INFO";
        console.log(`  ${icon} [${label}] pos ${String(issue.index).padStart(3, " ")}  ${issue.message}`);
    }
    const parts = [];
    if (warns.length > 0)
        parts.push(`${warns.length} предупреждений`);
    if (infos.length > 0)
        parts.push(`${infos.length} замечаний`);
    console.log(`\nlintChildren: ${parts.join(", ")}.`);
}
//# sourceMappingURL=lint.js.map