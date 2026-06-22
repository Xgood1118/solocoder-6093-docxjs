import { WordDocument } from './word-document';
import { DocumentParser } from './document-parser';
import { HtmlRenderer } from './html-renderer';
import { h } from './html';

let searchKeyHandler: ((e: KeyboardEvent) => void) | null = null;

export interface UserStyle {
    selector: string;
    properties: Record<string, string>;
}

export interface OutlineItem {
    id: string;
    text: string;
    level: number;
    element: HTMLElement;
    children: OutlineItem[];
}

export interface Options {
    inWrapper: boolean;
    hideWrapperOnPrint: boolean;
    ignoreWidth: boolean;
    ignoreHeight: boolean;
    ignoreFonts: boolean;
    breakPages: boolean;
    debug: boolean;
    experimental: boolean;
    className: string;
    trimXmlDeclaration: boolean;
    renderHeaders: boolean;
    renderFooters: boolean;
    renderFootnotes: boolean;
	renderEndnotes: boolean;
    ignoreLastRenderedPageBreak: boolean;
	useBase64URL: boolean;
	renderChanges: boolean;
    renderComments: boolean;
    renderAltChunks: boolean;
    h: typeof h;
    showNavigation: boolean;
    navigationDefaultLevel: number;
    userStyles: UserStyle[];
}

export const defaultOptions: Options = {
    ignoreHeight: false,
    ignoreWidth: false,
    ignoreFonts: false,
    breakPages: true,
    debug: false,
    experimental: false,
    className: "docx",
    inWrapper: true,
    hideWrapperOnPrint: false,
    trimXmlDeclaration: true,
    ignoreLastRenderedPageBreak: true,
    renderHeaders: true,
    renderFooters: true,
    renderFootnotes: true,
	renderEndnotes: true,
	useBase64URL: false,
	renderChanges: false,
    renderComments: false,
    renderAltChunks: true,
    h: h,
    showNavigation: true,
    navigationDefaultLevel: 3,
    userStyles: []
};

export function parseAsync(data: Blob | any, userOptions?: Partial<Options>): Promise<any>  {
    const ops = { ...defaultOptions, ...userOptions };
    return WordDocument.load(data, new DocumentParser(ops), ops);
}

export async function renderDocument(document: any, userOptions?: Partial<Options>): Promise<any> {
    const ops = { ...defaultOptions, ...userOptions };
    const renderer = new HtmlRenderer();
    return await renderer.render(document, ops);
}

export async function renderAsync(data: Blob | any, bodyContainer: HTMLElement, styleContainer?: HTMLElement, userOptions?: Partial<Options>): Promise<any> {
	const doc = await parseAsync(data, userOptions);
	const nodes = await renderDocument(doc, userOptions);
    const ops = { ...defaultOptions, ...userOptions };

    styleContainer ??= bodyContainer;
    styleContainer.innerHTML = "";
    bodyContainer.innerHTML = "";

    for (let n of nodes) {
        const c = n.nodeName === "STYLE" ? styleContainer : bodyContainer;
        c.appendChild(n);
    }

    cleanupNavigation(bodyContainer, ops);
    cleanupSearch(ops);

    injectUserStyles(styleContainer, ops);

    if (ops.showNavigation) {
        buildNavigation(bodyContainer, styleContainer, ops);
    }

    setupSearch(bodyContainer, styleContainer, ops);

    return doc;
}

function injectUserStyles(styleContainer: HTMLElement, options: Options): void {
    if (!options.userStyles || options.userStyles.length === 0) return;

    let cssText = "";
    for (const style of options.userStyles) {
        cssText += `${style.selector} {\r\n`;
        for (const [key, value] of Object.entries(style.properties)) {
            cssText += `  ${key}: ${value};\r\n`;
        }
        cssText += "}\r\n";
    }

    const styleEl = document.createElement("style");
    styleEl.textContent = cssText;
    styleEl.setAttribute("data-docx-user-styles", "true");
    styleContainer.appendChild(styleEl);
}

function cleanupNavigation(bodyContainer: HTMLElement, options: Options): void {
    const className = options.className;
    const parent = bodyContainer.parentNode;

    if (parent) {
        const navContainer = parent.querySelector(`.${className}-nav-container`);
        const navToggle = parent.querySelector(`.${className}-nav-toggle`);
        if (navContainer) navContainer.remove();
        if (navToggle) navToggle.remove();
    }

    bodyContainer.classList.remove(`${className}-body-with-nav`);
}

function cleanupSearch(options: Options): void {
    const className = options.className;
    const searchContainer = document.querySelector(`.${className}-search-container`);
    if (searchContainer) searchContainer.remove();

    if (searchKeyHandler) {
        document.removeEventListener("keydown", searchKeyHandler);
        searchKeyHandler = null;
    }
}

function buildNavigation(bodyContainer: HTMLElement, styleContainer: HTMLElement, options: Options): void {
    const className = options.className;
    const headings = collectHeadings(bodyContainer, className);

    if (headings.length === 0) return;

    const outlineTree = buildOutlineTree(headings);

    const navStyle = `
.${className}-nav-container {
    position: fixed;
    left: 0;
    top: 0;
    bottom: 0;
    width: 280px;
    background: #f8f9fa;
    border-right: 1px solid #dee2e6;
    overflow-y: auto;
    z-index: 1000;
    padding: 16px 0;
    box-sizing: border-box;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
}
.${className}-nav-header {
    padding: 0 16px 12px;
    font-weight: 600;
    font-size: 14px;
    color: #495057;
    border-bottom: 1px solid #dee2e6;
    margin-bottom: 8px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.${className}-nav-toggle {
    position: fixed;
    left: 280px;
    top: 12px;
    z-index: 1001;
    background: #f8f9fa;
    border: 1px solid #dee2e6;
    border-left: none;
    border-radius: 0 4px 4px 0;
    padding: 4px 8px;
    cursor: pointer;
    font-size: 12px;
    color: #495057;
}
.${className}-nav-toggle:hover {
    background: #e9ecef;
}
.${className}-nav-container.collapsed {
    width: 0;
    overflow: hidden;
    border-right: none;
}
.${className}-nav-container.collapsed + .${className}-nav-toggle {
    left: 0;
    border-left: 1px solid #dee2e6;
    border-radius: 0 4px 4px 0;
}
.${className}-nav-ul {
    list-style: none;
    padding: 0;
    margin: 0;
}
.${className}-nav-li {
    margin: 0;
}
.${className}-nav-item {
    display: flex;
    align-items: center;
    padding: 6px 16px;
    cursor: pointer;
    color: #495057;
    text-decoration: none;
    font-size: 13px;
    line-height: 1.4;
    border-left: 3px solid transparent;
    transition: background-color 0.15s, border-color 0.15s;
}
.${className}-nav-item:hover {
    background: #e9ecef;
}
.${className}-nav-item.active {
    background: #e7f1ff;
    border-left-color: #0d6efd;
    color: #0d6efd;
    font-weight: 500;
}
.${className}-nav-expand {
    display: inline-flex;
    align-items: center;
    justify-content: center;
    width: 16px;
    height: 16px;
    margin-right: 4px;
    cursor: pointer;
    flex-shrink: 0;
    font-size: 10px;
    color: #6c757d;
    user-select: none;
}
.${className}-nav-expand:hover {
    color: #495057;
}
.${className}-nav-expand.empty {
    visibility: hidden;
}
.${className}-nav-children {
    list-style: none;
    padding-left: 0;
    margin: 0;
    display: none;
}
.${className}-nav-children.expanded {
    display: block;
}
.${className}-nav-text {
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
    flex: 1;
}
.${className}-nav-heading-marker {
    display: inline-block;
    width: 20px;
    flex-shrink: 0;
    color: #adb5bd;
    font-size: 11px;
}
.${className}-body-with-nav {
    margin-left: 280px !important;
}
@media print {
    .${className}-nav-container, .${className}-nav-toggle {
        display: none !important;
    }
    .${className}-body-with-nav {
        margin-left: 0 !important;
    }
}
`;

    const navStyleEl = document.createElement("style");
    navStyleEl.textContent = navStyle;
    styleContainer.appendChild(navStyleEl);

    const navContainer = document.createElement("div");
    navContainer.className = `${className}-nav-container`;
    navContainer.setAttribute("data-docx-nav", "true");

    const navHeader = document.createElement("div");
    navHeader.className = `${className}-nav-header`;
    navHeader.textContent = "目录";
    navContainer.appendChild(navHeader);

    const navList = document.createElement("ul");
    navList.className = `${className}-nav-ul`;

    renderOutlineItems(outlineTree, navList, className, options.navigationDefaultLevel);

    navContainer.appendChild(navList);

    const toggleBtn = document.createElement("button");
    toggleBtn.className = `${className}-nav-toggle`;
    toggleBtn.textContent = "◀";
    toggleBtn.title = "隐藏/显示目录";
    toggleBtn.addEventListener("click", () => {
        navContainer.classList.toggle("collapsed");
        bodyContainer.classList.toggle(`${className}-body-with-nav`);
        toggleBtn.textContent = navContainer.classList.contains("collapsed") ? "▶" : "◀";
    });

    bodyContainer.parentNode?.insertBefore(navContainer, bodyContainer);
    bodyContainer.parentNode?.insertBefore(toggleBtn, bodyContainer.nextSibling);
    bodyContainer.classList.add(`${className}-body-with-nav`);

    setupScrollSpy(bodyContainer, headings, className);
}

function collectHeadings(container: HTMLElement, className: string): { element: HTMLElement; level: number; text: string; id: string }[] {
    const headings: { element: HTMLElement; level: number; text: string; id: string }[] = [];
    const paragraphs = container.querySelectorAll('p');
    let idCounter = 0;

    paragraphs.forEach((p) => {
        let level: number | null = null;

        const outlineLevelAttr = p.getAttribute('data-outline-level');
        if (outlineLevelAttr) {
            const parsedLevel = parseInt(outlineLevelAttr, 10);
            if (!isNaN(parsedLevel) && parsedLevel >= 1 && parsedLevel <= 9) {
                level = parsedLevel;
            }
        }

        if (level === null) {
            const classList = Array.from(p.classList);
            for (const cls of classList) {
                if (cls.startsWith(`${className}_`)) {
                    const styleId = cls.substring(className.length + 1);
                    if (/^heading\d+$/i.test(styleId)) {
                        const levelMatch = styleId.match(/heading(\d+)/i);
                        if (levelMatch) {
                            level = parseInt(levelMatch[1], 10);
                            break;
                        }
                    }
                }
            }
        }

        if (level !== null && level >= 1 && level <= 9) {
            const id = `docx-heading-${idCounter++}`;
            p.id = id;
            headings.push({
                element: p as HTMLElement,
                level,
                text: p.textContent?.trim() || '',
                id
            });
        }
    });

    return headings;
}

function buildOutlineTree(headings: { element: HTMLElement; level: number; text: string; id: string }[]): OutlineItem[] {
    const root: OutlineItem[] = [];
    const stack: OutlineItem[] = [];

    for (const heading of headings) {
        const item: OutlineItem = {
            id: heading.id,
            text: heading.text,
            level: heading.level,
            element: heading.element,
            children: []
        };

        while (stack.length > 0 && stack[stack.length - 1].level >= heading.level) {
            stack.pop();
        }

        if (stack.length === 0) {
            root.push(item);
        } else {
            stack[stack.length - 1].children.push(item);
        }

        stack.push(item);
    }

    return root;
}

function renderOutlineItems(items: OutlineItem[], parent: HTMLElement, className: string, defaultLevel: number, currentLevel: number = 1): void {
    for (const item of items) {
        const li = document.createElement("li");
        li.className = `${className}-nav-li`;
        li.setAttribute("data-level", item.level.toString());

        const itemDiv = document.createElement("div");
        itemDiv.className = `${className}-nav-item`;
        itemDiv.setAttribute("data-target", item.id);

        const expandBtn = document.createElement("span");
        expandBtn.className = `${className}-nav-expand`;
        expandBtn.textContent = item.children.length > 0 ? "▶" : "";
        if (item.children.length === 0) {
            expandBtn.classList.add("empty");
        }

        const headingMarker = document.createElement("span");
        headingMarker.className = `${className}-nav-heading-marker`;
        headingMarker.textContent = `H${item.level}`;

        const textSpan = document.createElement("span");
        textSpan.className = `${className}-nav-text`;
        textSpan.textContent = item.text;

        itemDiv.appendChild(expandBtn);
        itemDiv.appendChild(headingMarker);
        itemDiv.appendChild(textSpan);

        itemDiv.addEventListener("click", (e) => {
            if ((e.target as HTMLElement).classList.contains(`${className}-nav-expand`)) return;
            const target = document.getElementById(item.id);
            if (target) {
                target.scrollIntoView({ behavior: "smooth", block: "start" });
            }
        });

        expandBtn.addEventListener("click", (e) => {
            e.stopPropagation();
            const childrenUl = li.querySelector(`.${className}-nav-children`);
            if (childrenUl) {
                childrenUl.classList.toggle("expanded");
                expandBtn.textContent = childrenUl.classList.contains("expanded") ? "▼" : "▶";
            }
        });

        li.appendChild(itemDiv);

        if (item.children.length > 0) {
            const childrenUl = document.createElement("ul");
            childrenUl.className = `${className}-nav-children`;
            if (item.level < defaultLevel) {
                childrenUl.classList.add("expanded");
                expandBtn.textContent = "▼";
            }
            renderOutlineItems(item.children, childrenUl, className, defaultLevel, currentLevel + 1);
            li.appendChild(childrenUl);
        }

        parent.appendChild(li);
    }
}

function setupScrollSpy(container: HTMLElement, headings: { element: HTMLElement; level: number; text: string; id: string }[], className: string): void {
    const navItems = container.parentElement?.querySelectorAll(`.${className}-nav-item`);
    if (!navItems || navItems.length === 0) return;

    let ticking = false;

    function updateActive(): void {
        const scrollTop = container.scrollTop;
        const offset = 100;

        let currentId: string | null = null;
        for (let i = headings.length - 1; i >= 0; i--) {
            const heading = headings[i];
            const rect = heading.element.getBoundingClientRect();
            const containerRect = container.getBoundingClientRect();
            const relativeTop = rect.top - containerRect.top;

            if (relativeTop <= offset) {
                currentId = heading.id;
                break;
            }
        }

        if (headings.length > 0 && !currentId) {
            const firstHeading = headings[0];
            const rect = firstHeading.element.getBoundingClientRect();
            const containerRect = container.getBoundingClientRect();
            if (rect.top - containerRect.top > offset) {
                currentId = firstHeading.id;
            }
        }

        navItems.forEach((item) => {
            item.classList.remove("active");
            if (item.getAttribute("data-target") === currentId) {
                item.classList.add("active");
                const navContainer = container.parentElement?.querySelector(`.${className}-nav-container`);
                if (navContainer && item.getBoundingClientRect) {
                    const itemRect = item.getBoundingClientRect();
                    const navRect = navContainer.getBoundingClientRect();
                    if (itemRect.top < navRect.top || itemRect.bottom > navRect.bottom) {
                        item.scrollIntoView({ block: "nearest" });
                    }
                }
            }
        });

        ticking = false;
    }

    container.addEventListener("scroll", () => {
        if (!ticking) {
            requestAnimationFrame(() => {
                updateActive();
            });
            ticking = true;
        }
    });

    setTimeout(updateActive, 100);
}

function setupSearch(bodyContainer: HTMLElement, styleContainer: HTMLElement, options: Options): void {
    const className = options.className;

    const searchStyle = `
.${className}-search-container {
    position: fixed;
    top: 12px;
    right: 12px;
    z-index: 1002;
    background: white;
    border: 1px solid #dee2e6;
    border-radius: 6px;
    box-shadow: 0 2px 8px rgba(0, 0, 0, 0.15);
    padding: 8px;
    display: none;
    font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
    font-size: 13px;
}
.${className}-search-container.visible {
    display: block;
}
.${className}-search-input-row {
    display: flex;
    gap: 6px;
    align-items: center;
}
.${className}-search-input {
    width: 200px;
    padding: 6px 8px;
    border: 1px solid #ced4da;
    border-radius: 4px;
    font-size: 13px;
    outline: none;
}
.${className}-search-input:focus {
    border-color: #0d6efd;
    box-shadow: 0 0 0 2px rgba(13, 110, 253, 0.25);
}
.${className}-search-btn {
    padding: 6px 10px;
    border: 1px solid #ced4da;
    background: #f8f9fa;
    border-radius: 4px;
    cursor: pointer;
    font-size: 12px;
    color: #495057;
}
.${className}-search-btn:hover {
    background: #e9ecef;
}
.${className}-search-btn.primary {
    background: #0d6efd;
    color: white;
    border-color: #0d6efd;
}
.${className}-search-btn.primary:hover {
    background: #0b5ed7;
}
.${className}-search-options {
    display: flex;
    align-items: center;
    gap: 8px;
    margin-top: 8px;
    padding-top: 8px;
    border-top: 1px solid #f1f3f5;
}
.${className}-search-checkbox {
    display: flex;
    align-items: center;
    gap: 4px;
    cursor: pointer;
    color: #6c757d;
    font-size: 12px;
}
.${className}-search-count {
    color: #6c757d;
    font-size: 12px;
    margin-left: auto;
}
.${className}-search-highlight {
    background-color: #fff3cd !important;
    padding: 1px 0;
}
.${className}-search-highlight.current {
    background-color: #fd7e14 !important;
    color: white;
}
@media print {
    .${className}-search-container {
        display: none !important;
    }
}
`;

    const searchStyleEl = document.createElement("style");
    searchStyleEl.textContent = searchStyle;
    styleContainer.appendChild(searchStyleEl);

    const searchContainer = document.createElement("div");
    searchContainer.className = `${className}-search-container`;

    const inputRow = document.createElement("div");
    inputRow.className = `${className}-search-input-row`;

    const searchInput = document.createElement("input");
    searchInput.type = "text";
    searchInput.className = `${className}-search-input`;
    searchInput.placeholder = "搜索文档...";

    const prevBtn = document.createElement("button");
    prevBtn.className = `${className}-search-btn`;
    prevBtn.textContent = "上一个";
    prevBtn.title = "上一个匹配项 (Shift+Enter)";

    const nextBtn = document.createElement("button");
    nextBtn.className = `${className}-search-btn primary`;
    nextBtn.textContent = "下一个";
    nextBtn.title = "下一个匹配项 (Enter)";

    const closeBtn = document.createElement("button");
    closeBtn.className = `${className}-search-btn`;
    closeBtn.textContent = "✕";
    closeBtn.title = "关闭 (Esc)";

    inputRow.appendChild(searchInput);
    inputRow.appendChild(prevBtn);
    inputRow.appendChild(nextBtn);
    inputRow.appendChild(closeBtn);

    const optionsRow = document.createElement("div");
    optionsRow.className = `${className}-search-options`;

    const caseLabel = document.createElement("label");
    caseLabel.className = `${className}-search-checkbox`;
    const caseCheckbox = document.createElement("input");
    caseCheckbox.type = "checkbox";
    caseLabel.appendChild(caseCheckbox);
    caseLabel.appendChild(document.createTextNode("区分大小写"));

    const countSpan = document.createElement("span");
    countSpan.className = `${className}-search-count`;
    countSpan.textContent = "0 个匹配";

    optionsRow.appendChild(caseLabel);
    optionsRow.appendChild(countSpan);

    searchContainer.appendChild(inputRow);
    searchContainer.appendChild(optionsRow);

    document.body.appendChild(searchContainer);

    let searchMatches: HTMLElement[] = [];
    let currentMatchIndex = -1;
    let searchQuery = "";
    let caseSensitive = false;

    function clearHighlights(): void {
        const highlights = bodyContainer.querySelectorAll(`.${className}-search-highlight`);
        highlights.forEach((hl) => {
            const parent = hl.parentNode;
            if (parent) {
                while (hl.firstChild) {
                    parent.insertBefore(hl.firstChild, hl);
                }
                parent.removeChild(hl);
                parent.normalize();
            }
        });
        searchMatches = [];
        currentMatchIndex = -1;
    }

    function performSearch(query: string): void {
        clearHighlights();
        searchQuery = query;

        if (!query) {
            countSpan.textContent = "0 个匹配";
            return;
        }

        const contentElements = bodyContainer.querySelectorAll(`p, td, th, li`);
        const flags = caseSensitive ? "g" : "gi";
        const regex = new RegExp(`(${escapeRegExp(query)})`, flags);

        contentElements.forEach((el) => {
            if (el.closest(`.${className}-search-container`)) return;
            highlightTextInElement(el as HTMLElement, regex, className);
        });

        searchMatches = Array.from(bodyContainer.querySelectorAll(`.${className}-search-highlight`));
        countSpan.textContent = `${searchMatches.length} 个匹配`;

        if (searchMatches.length > 0) {
            currentMatchIndex = 0;
            highlightCurrentMatch();
        }
    }

    function highlightTextInElement(element: HTMLElement, regex: RegExp, className: string): void {
        const textNodes: Text[] = [];
        const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT, null);
        let node: Node | null;
        while ((node = walker.nextNode())) {
            if (node.textContent && node.textContent.trim()) {
                textNodes.push(node as Text);
            }
        }

        textNodes.forEach((textNode) => {
            const text = textNode.textContent || "";
            if (!regex.test(text)) return;
            regex.lastIndex = 0;

            const parent = textNode.parentNode;
            if (!parent) return;

            const fragment = document.createDocumentFragment();
            let lastIndex = 0;
            let match: RegExpExecArray | null;

            while ((match = regex.exec(text)) !== null) {
                if (match.index > lastIndex) {
                    fragment.appendChild(document.createTextNode(text.substring(lastIndex, match.index)));
                }

                const span = document.createElement("span");
                span.className = `${className}-search-highlight`;
                span.textContent = match[0];
                fragment.appendChild(span);

                lastIndex = match.index + match[0].length;

                if (match.index === regex.lastIndex) {
                    regex.lastIndex++;
                }
            }

            if (lastIndex < text.length) {
                fragment.appendChild(document.createTextNode(text.substring(lastIndex)));
            }

            parent.replaceChild(fragment, textNode);
        });
    }

    function escapeRegExp(string: string): string {
        return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    }

    function highlightCurrentMatch(): void {
        const highlights = bodyContainer.querySelectorAll(`.${className}-search-highlight`);
        highlights.forEach((hl) => hl.classList.remove("current"));

        if (currentMatchIndex >= 0 && currentMatchIndex < searchMatches.length) {
            const current = searchMatches[currentMatchIndex];
            current.classList.add("current");
            current.scrollIntoView({ behavior: "smooth", block: "center" });
        }
    }

    function goToNext(): void {
        if (searchMatches.length === 0) return;
        currentMatchIndex = (currentMatchIndex + 1) % searchMatches.length;
        highlightCurrentMatch();
    }

    function goToPrev(): void {
        if (searchMatches.length === 0) return;
        currentMatchIndex = (currentMatchIndex - 1 + searchMatches.length) % searchMatches.length;
        highlightCurrentMatch();
    }

    function openSearch(): void {
        searchContainer.classList.add("visible");
        searchInput.focus();
        searchInput.select();
    }

    function closeSearch(): void {
        searchContainer.classList.remove("visible");
        clearHighlights();
        searchInput.value = "";
        countSpan.textContent = "0 个匹配";
    }

    searchInput.addEventListener("input", () => {
        performSearch(searchInput.value);
    });

    searchInput.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
            e.preventDefault();
            if (e.shiftKey) {
                goToPrev();
            } else {
                goToNext();
            }
        } else if (e.key === "Escape") {
            e.preventDefault();
            closeSearch();
        }
    });

    nextBtn.addEventListener("click", goToNext);
    prevBtn.addEventListener("click", goToPrev);
    closeBtn.addEventListener("click", closeSearch);

    caseCheckbox.addEventListener("change", () => {
        caseSensitive = caseCheckbox.checked;
        performSearch(searchInput.value);
    });

    searchKeyHandler = (e: KeyboardEvent) => {
        if ((e.ctrlKey || e.metaKey) && e.key === "f") {
            e.preventDefault();
            openSearch();
        } else if (e.key === "Escape" && searchContainer.classList.contains("visible")) {
            closeSearch();
        }
    };

    document.addEventListener("keydown", searchKeyHandler);
}

export function exportText(container: HTMLElement, options?: Partial<Options>): string {
    const ops = { ...defaultOptions, ...options };
    const className = ops.className;
    let result = "";

    const sections = container.querySelectorAll(`section.${className}`);
    if (sections.length === 0) {
        result = extractTextFromNode(container, className);
    } else {
        sections.forEach((section, sectionIndex) => {
            if (sectionIndex > 0) {
                result += "\n\n";
            }
            result += extractTextFromNode(section as HTMLElement, className);
        });
    }

    return result.trim();
}

function extractTextFromNode(node: HTMLElement, className: string): string {
    let text = "";
    const walker = document.createTreeWalker(node, NodeFilter.SHOW_ELEMENT, {
        acceptNode: (el: Node) => {
            if (el.nodeType === 1) {
                const tag = (el as Element).tagName.toLowerCase();
                if (['p', 'table', 'tr', 'td', 'th', 'li', 'ul', 'ol', 'br'].includes(tag)) {
                    return NodeFilter.FILTER_ACCEPT;
                }
            }
            return NodeFilter.FILTER_SKIP;
        }
    });

    let current: Node | null = walker.nextNode();
    let lastWasParagraph = false;

    while (current) {
        const el = current as HTMLElement;
        const tag = el.tagName.toLowerCase();

        if (tag === 'p') {
            const paragraphText = extractParagraphText(el, className);
            if (paragraphText) {
                if (text && !text.endsWith('\n')) {
                    text += '\n';
                }
                text += paragraphText;
                lastWasParagraph = true;
            }
        } else if (tag === 'table') {
            if (text && !text.endsWith('\n')) {
                text += '\n';
            }
            text += extractTableText(el);
            lastWasParagraph = false;
        } else if (tag === 'li') {
            if (text && !text.endsWith('\n')) {
                text += '\n';
            }
            text += el.textContent?.trim() || '';
            lastWasParagraph = false;
        } else if (tag === 'br') {
            text += '\n';
            lastWasParagraph = false;
        }

        current = walker.nextNode();
    }

    return text;
}

function extractParagraphText(paragraph: HTMLElement, className: string): string {
    const text = paragraph.textContent?.trim() || '';
    if (!text) return '';

    let level: number | null = null;

    const outlineLevelAttr = paragraph.getAttribute('data-outline-level');
    if (outlineLevelAttr) {
        const parsedLevel = parseInt(outlineLevelAttr, 10);
        if (!isNaN(parsedLevel) && parsedLevel >= 1 && parsedLevel <= 9) {
            level = parsedLevel;
        }
    }

    if (level === null) {
        const classList = Array.from(paragraph.classList);
        for (const cls of classList) {
            if (cls.startsWith(`${className}_`)) {
                const styleId = cls.substring(className.length + 1);
                if (/^heading\d+$/i.test(styleId)) {
                    const levelMatch = styleId.match(/heading(\d+)/i);
                    if (levelMatch) {
                        level = parseInt(levelMatch[1], 10);
                        break;
                    }
                }
            }
        }
    }

    if (level !== null && level >= 1 && level <= 9) {
        return '#'.repeat(level) + ' ' + text;
    }

    return text;
}

function extractTableText(table: HTMLElement): string {
    let result = "";
    const rows = table.querySelectorAll('tr');

    rows.forEach((row, rowIndex) => {
        const cells = row.querySelectorAll('td, th');
        const cellTexts: string[] = [];

        cells.forEach((cell) => {
            cellTexts.push(cell.textContent?.trim() || '');
        });

        if (cellTexts.length > 0) {
            result += cellTexts.join('\t');
            if (rowIndex < rows.length - 1) {
                result += '\n';
            }
        }
    });

    return result;
}