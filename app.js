/* ==========================================================================
   MD2MSWord — Core Application Engine
   ==========================================================================
   Optimized for Microsoft Word Desktop Software
   - Extracts LaTeX Math & ChatGPT/Claude/Gemini/DeepSeek formats
   - Generates pure MathML (<math>) for MS Word equation engine
   - Strips KaTeX visual DOM to prevent character doubling in MS Word
   ========================================================================== */

(function () {
    'use strict';

    // DOM Elements
    const markdownInput = document.getElementById('markdownInput');
    const wordPreview = document.getElementById('wordPreview');
    const copyWordBtn = document.getElementById('copyWordBtn');
    const clearBtn = document.getElementById('clearBtn');
    const fontSelect = document.getElementById('fontSelect');
    const headingColorSelect = document.getElementById('headingColorSelect');
    const colorSwatch = document.getElementById('colorSwatch');
    const toastNotification = document.getElementById('toastNotification');

    // Configure Marked Parser
    marked.use({
        breaks: true,
        gfm: true
    });

    // =========================================================================
    // Math Extraction & Protection Engine
    // =========================================================================

    /**
     * Protects LaTeX math delimiters (including ChatGPT/Claude/Gemini formats)
     * from being corrupted by marked.js markdown parsing.
     */
    function extractMath(text) {
        const mathBlocks = [];
        let index = 0;

        // 1. Protect code blocks (fenced ``` and inline `)
        const codeSlots = [];
        text = text.replace(/```[\s\S]*?```/g, (match) => {
            const slot = `\x00CBLK_${codeSlots.length}\x00`;
            codeSlots.push({ slot, code: match });
            return slot;
        });
        text = text.replace(/`[^`\n]+`/g, (match) => {
            const slot = `\x00CINL_${codeSlots.length}\x00`;
            codeSlots.push({ slot, code: match });
            return slot;
        });

        // 2. Standard Display Math ($$ ... $$ or \[ ... \])
        text = text.replace(/\$\$([\s\S]+?)\$\$/g, (_, tex) => {
            const ph = `\x00DMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: true, ph });
            return `\n\n${ph}\n\n`;
        });
        text = text.replace(/\\\[([\s\S]+?)\\\]/g, (_, tex) => {
            const ph = `\x00DMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: true, ph });
            return `\n\n${ph}\n\n`;
        });

        // 3. ChatGPT / Markdown Unescaped Bracket Display Math:
        // Matches [ \n content \n ] or [ content ] formatted as block math
        // Excludes markdown links [label](url)
        text = text.replace(/(?<=^|\n)\s*\[\s*\n?([\s\S]+?)\n?\s*\](?=\n|$)(?!\()/g, (_, tex) => {
            const ph = `\x00DMATH_${index}\x00`;
            let cleanedTex = tex.trim().replace(/^=+\s*$/gm, '');
            mathBlocks.push({ index: index++, tex: cleanedTex.trim(), display: true, ph });
            return `\n\n${ph}\n\n`;
        });

        // Single-line brackets containing TeX commands e.g. [ \left| \frac{v_o}{v_{in}} \right| = ... ]
        text = text.replace(/\[\s*([^\n\]]*?\\[a-zA-Z]+[^\n\]]*?)\s*\](?!\()/g, (_, tex) => {
            const ph = `\x00DMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: true, ph });
            return `\n\n${ph}\n\n`;
        });

        // 4. Standard Inline Math ($ ... $ or \( ... \))
        text = text.replace(/(?<!\$)\$(?!\$)(\S(?:[^\$\n]*?\S)?)\$(?!\$)/g, (_, tex) => {
            const ph = `\x00IMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: false, ph });
            return ph;
        });
        text = text.replace(/\\\((.+?)\\\)/g, (_, tex) => {
            const ph = `\x00IMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: false, ph });
            return ph;
        });

        // 5. ChatGPT Parenthesis Inline Math with balanced parens:
        // Matches ( \tan^{-1}(x) ), ( \frac{f}{f_H} \approx 0 ), ( f \ll f_H )
        text = text.replace(/\(\s*(((?:[^\(\)\n]|\([^\)\n]*\))+?\\[a-zA-Z]+(?:[^\(\)\n]|\([^\)\n]*\))*?))\s*\)/g, (_, tex) => {
            const ph = `\x00IMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: false, ph });
            return ph;
        });

        // Matches equations like ( a = 1 ), ( b = 2 ), ( f = f_H ), ( 1 + jx )
        text = text.replace(/\(\s*([a-zA-Z0-9_\^\{\}\s\+\-\*\/\=]+\s*=\s*[a-zA-Z0-9_\^\{\}\s\+\-\*\/\=]+)\s*\)/g, (_, tex) => {
            const ph = `\x00IMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: false, ph });
            return ph;
        });

        text = text.replace(/\(\s*([0-9]+\s*[\+\-\*\/]\s*[a-zA-Z0-9_\^]+|[a-zA-Z]\s*[\+\-\*\/]\s*[a-zA-Z0-9_\^]+)\s*\)/g, (_, tex) => {
            const ph = `\x00IMATH_${index}\x00`;
            mathBlocks.push({ index: index++, tex: tex.trim(), display: false, ph });
            return ph;
        });

        // 6. Restore Code Blocks
        codeSlots.forEach(({ slot, code }) => {
            text = text.replace(slot, code);
        });

        return { text, mathBlocks };
    }

    // =========================================================================
    // Word Desktop HTML Enhancement Pipeline
    // =========================================================================

    /**
     * Enhances HTML formatting specifically for MS Word Desktop compatibility:
     * - Transforms GFM task list checkboxes (- [x] and - [ ]) into Word symbols
     * - Transforms GitHub Alert Callouts ([!NOTE], [!TIP], [!IMPORTANT], [!WARNING], [!CAUTION]) into Word Callout Boxes
     * - Transforms <del>/<s> strikethrough, <mark> highlight, <kbd> key tags into Office inline styles
     * - Transforms <details><summary> accordions into Office document panels
     */
    function enhanceHtmlForWord(html) {
        // 1. Task List Checkboxes (- [x] and - [ ])
        html = html.replace(/<li[^>]*>\s*<input[^>]*checked[^>]*>\s*([\s\S]*?)<\/li>/gi, (match, content) => {
            return `<li style="list-style-type: none; margin-left: -15pt;"><span style="font-family: 'Segoe UI Symbol', 'Arial'; font-weight: bold; color: #107C41; font-size: 11pt;">☑</span>&nbsp; <strong>${content.trim()}</strong></li>`;
        });
        html = html.replace(/<li[^>]*>\s*<input[^>]*>\s*([\s\S]*?)<\/li>/gi, (match, content) => {
            return `<li style="list-style-type: none; margin-left: -15pt;"><span style="font-family: 'Segoe UI Symbol', 'Arial'; color: #595959; font-size: 11pt;">☐</span>&nbsp; ${content.trim()}</li>`;
        });

        // 2. GitHub Callout Alerts ([!NOTE], [!TIP], [!IMPORTANT], [!WARNING], [!CAUTION])
        const alertTypes = {
            'NOTE': { color: '#0969DA', bg: '#F0F6FF', title: 'ℹ️ NOTE' },
            'TIP': { color: '#1A7F37', bg: '#F0FDF4', title: '💡 TIP' },
            'IMPORTANT': { color: '#8250DF', bg: '#F8F5FF', title: '🟣 IMPORTANT' },
            'WARNING': { color: '#9A6700', bg: '#FFF8E6', title: '⚠️ WARNING' },
            'CAUTION': { color: '#CF222E', bg: '#FFEBE9', title: '🚨 CAUTION' }
        };

        html = html.replace(/blockquote>\s*<p>\s*\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]\s*([\s\S]*?)<\/blockquote/gi, (_, type, body) => {
            const config = alertTypes[type.toUpperCase()] || alertTypes['NOTE'];
            return `table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 10pt 0; width: 100%;">\n<tr>\n  <td style="border-left: 4pt solid ${config.color}; background-color: ${config.bg}; padding: 8pt 12pt; color: #24292F; font-size: 11pt; text-align: left !important; text-align-last: left !important; word-spacing: normal !important; letter-spacing: normal !important;">\n    <p style="font-weight: bold; color: ${config.color}; margin: 0 0 4pt 0; font-size: 10.5pt; text-align: left !important; text-align-last: left !important; word-spacing: normal !important; letter-spacing: normal !important;">${config.title}</p>\n    <div style="color: #24292F; font-size: 11pt; text-align: left !important; text-align-last: left !important; word-spacing: normal !important; letter-spacing: normal !important;">${body.trim()}</div>\n  </td>\n</tr>\n</table`;
        });

        // Fallback inline alert titles
        html = html.replace(/\[!(NOTE|TIP|IMPORTANT|WARNING|CAUTION)\]/gi, (match, type) => {
            const config = alertTypes[type.toUpperCase()];
            return config ? `<strong style="color: ${config.color};">${config.title}</strong>` : match;
        });

        // 3. Strikethrough & Highlight & Kbd
        html = html.replace(/<del[^>]*>([\s\S]*?)<\/del>/gi, '<span style="text-decoration: line-through; color: #595959;">$1</span>');
        html = html.replace(/<s[^>]*>([\s\S]*?)<\/s>/gi, '<span style="text-decoration: line-through; color: #595959;">$1</span>');
        html = html.replace(/<mark[^>]*>([\s\S]*?)<\/mark>/gi, '<span style="background-color: #FFF200; color: #000000; padding: 1pt 3pt; border-radius: 2pt;">$1</span>');
        html = html.replace(/<kbd[^>]*>([\s\S]*?)<\/kbd>/gi, '<code style="font-family: \'Consolas\', monospace; background-color: #F2F2F2; border: 1pt solid #7F7F7F; padding: 1pt 4pt; border-radius: 2pt; font-size: 9.5pt; color: #000000;">$1</code>');

        // 4. Accordion Details / Summary
        html = html.replace(/<details[^>]*>\s*<summary[^>]*>([\s\S]*?)<\/summary>([\s\S]*?)<\/details>/gi, (_, title, content) => {
            return `<table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 10pt 0; width: 100%;">\n<tr>\n  <td style="border: 1pt solid #BFBFBF; background-color: #F8F9FA; padding: 8pt 12pt; border-radius: 3pt; text-align: left !important; text-align-last: left !important; word-spacing: normal !important;">\n    <p style="font-weight: bold; color: #1F4E79; margin: 0 0 6pt 0; font-size: 11pt; text-align: left !important;">▶ ${title.trim()}</p>\n    <div style="text-align: left !important;">${content.trim()}</div>\n  </td>\n</tr>\n</table>`;
        });

        // 5. Depth-Aware Nested Blockquote Office Tables
        const headingColor = headingColorSelect.value || '#262626';
        const bqThemes = [
            { border: headingColor, bg: '#F0F4F8' },
            { border: '#2F5597', bg: '#E6ECF5' },
            { border: '#005F73', bg: '#E0F2F1' },
            { border: '#046A38', bg: '#E8F5E9' }
        ];
        let bqDepth = 0;
        html = html.replace(/<\/?blockquote[^>]*>/gi, (match) => {
            if (match.startsWith('</')) {
                bqDepth = Math.max(0, bqDepth - 1);
                return '</td></tr></table>';
            } else {
                const theme = bqThemes[bqDepth % bqThemes.length];
                bqDepth++;
                return `<table border="0" cellspacing="0" cellpadding="0" style="border-collapse: collapse; margin: 6pt 0; width: 100%;"><tr><td style="border-left: 4pt solid ${theme.border}; background-color: ${theme.bg}; padding: 6pt 10pt 6pt 12pt; color: #262626; font-style: italic; text-align: left !important; text-align-last: left !important; word-spacing: normal !important; letter-spacing: normal !important;">`;
            }
        });

        return html;
    }

    // =========================================================================
    // Rendering Pipeline (Web Preview)
    // =========================================================================

    function renderWebPreview() {
        const raw = markdownInput.value;

        if (!raw.trim()) {
            wordPreview.innerHTML = `
                <div class="empty-preview">
                    <div class="empty-preview-icon">📄</div>
                    <div class="empty-preview-title">Your rendered document preview will appear here</div>
                    <p style="font-size:13px;margin-top:6px;">Paste Markdown with LaTeX math on the left panel.</p>
                </div>`;
            return;
        }

        // 1. Extract Math
        const { text, mathBlocks } = extractMath(raw);

        // 2. Render Markdown to HTML & Enhance for Word
        let html = marked.parse(text);
        html = enhanceHtmlForWord(html);

        // Clean inline style attributes that force dark mode colors / backgrounds from pasted web HTML
        html = html.replace(/\s*style="[^"]*"/gi, (styleAttr) => {
            return styleAttr
                .replace(/background(-color)?\s*:[^;]+;?/gi, '')
                .replace(/color\s*:[^;]+;?/gi, '');
        });

        // 3. Inject Web Math (KaTeX with htmlAndMathml output)
        mathBlocks.forEach(({ tex, display, ph }) => {
            let rendered;
            try {
                rendered = katex.renderToString(tex, {
                    displayMode: display,
                    output: 'htmlAndMathml',
                    throwOnError: false
                });
            } catch (err) {
                rendered = display
                    ? `<div class="katex-error">$$${escapeHtml(tex)}$$</div>`
                    : `<span class="katex-error">$${escapeHtml(tex)}$</span>`;
            }

            if (display) {
                const wrapper = `<div class="katex-display">${rendered}</div>`;
                const wrappedPh = `<p>${ph}</p>`;
                if (html.includes(wrappedPh)) {
                    html = html.replace(wrappedPh, wrapper);
                } else {
                    html = html.replace(ph, wrapper);
                }
            } else {
                html = html.replace(ph, rendered);
            }
        });

        // 4. Clean structural wrapper tags & loose fragment comments
        html = html
            .replace(/<\/?(html|body|head|meta)[^>]*>/gi, '')
            .replace(/<!--\s*StartFragment\s*-->|<!--\s*EndFragment\s*-->/gi, '');

        // 5. Safely parse via DOMParser inside an isolated container to prevent loose </div> tags from closing wordPreview
        const parser = new DOMParser();
        const parsedDoc = parser.parseFromString(`<div>${html}</div>`, 'text/html');
        const safeContainer = parsedDoc.body.firstElementChild;

        // Populate wordPreview without container leakage
        wordPreview.replaceChildren(...safeContainer.childNodes);

        // Apply selected font style & heading accent color to preview
        const selectedFont = fontSelect.value || 'Times New Roman';
        const selectedHeadingColor = headingColorSelect.value || '#262626';

        wordPreview.style.fontFamily = `'${selectedFont}', sans-serif`;

        if (colorSwatch) {
            colorSwatch.style.backgroundColor = selectedHeadingColor;
        }

        // Apply heading accent color ONLY to h1, h2, h3 elements
        wordPreview.querySelectorAll('h1, h2, h3').forEach(heading => {
            heading.style.color = selectedHeadingColor;
        });
    }

    // =========================================================================
    // MS Word Desktop HTML Generator (MathML Engine)
    // =========================================================================

    /**
     * Builds clean HTML specifically formatted for MS Word Desktop.
     * Injects pure MathML (<math xmlns="http://www.w3.org/1998/Math/MathML">)
     * without KaTeX visual DOM nodes to avoid character doubling when pasting into Word.
     */
    function buildWordHtmlPayload(rawMarkdown) {
        const targetFont = fontSelect.value || 'Times New Roman';
        const headingColor = headingColorSelect.value || '#262626';

        // 1. Extract Math
        const { text, mathBlocks } = extractMath(rawMarkdown);

        // 2. Render Markdown to base HTML & Enhance for Word
        let html = marked.parse(text);
        html = enhanceHtmlForWord(html);

        // 3. Inject MathML for MS Word
        mathBlocks.forEach(({ tex, display, ph }) => {
            let mathmlTag = '';
            try {
                // Generate raw MathML string via KaTeX
                const katexMathML = katex.renderToString(tex, {
                    displayMode: display,
                    output: 'mathml',
                    throwOnError: false
                });

                // Extract purely the <math ...> ... </math> tag from KaTeX output
                const mathMatch = katexMathML.match(/<math[\s\S]*?<\/math>/i);
                if (mathMatch) {
                    mathmlTag = mathMatch[0];
                    // Ensure display attribute is explicitly set on <math>
                    if (display) {
                        mathmlTag = mathmlTag.replace(/<math/i, '<math display="block"');
                    } else {
                        mathmlTag = mathmlTag.replace(/<math/i, '<math display="inline"');
                    }

                    // Clean empty base elements in MathML that cause MS Word Desktop to render dotted square placeholders (⬚)
                    mathmlTag = mathmlTag
                        .replace(/<(msubsup|msub|msup|mmultiscripts)>\s*<mrow\s*\/>/gi, '<$1><mtext>&#x200B;</mtext>')
                        .replace(/<(msubsup|msub|msup|mmultiscripts)>\s*<mrow>\s*<\/mrow>/gi, '<$1><mtext>&#x200B;</mtext>')
                        .replace(/<(msubsup|msub|msup|mmultiscripts)>\s*<\/mrow>/gi, '<$1><mtext>&#x200B;</mtext>');
                } else {
                    mathmlTag = `<span>${escapeHtml(tex)}</span>`;
                }
            } catch (err) {
                mathmlTag = `<span>${escapeHtml(tex)}</span>`;
            }

            if (display) {
                const wrapper = `<p style="text-align:center;margin:12pt 0 12pt 0;font-family:'${targetFont}',sans-serif;">${mathmlTag}</p>`;
                const wrappedPh = `<p>${ph}</p>`;
                if (html.includes(wrappedPh)) {
                    html = html.replace(wrappedPh, wrapper);
                } else {
                    html = html.replace(ph, wrapper);
                }
            } else {
                html = html.replace(ph, mathmlTag);
            }
        });

        // 4. Wrap with Word Document HTML Headers & Office Namespaces
        const wordOfficeHtml = `<!DOCTYPE html>
<html xmlns:v="urn:schemas-microsoft-com:vml"
      xmlns:o="urn:schemas-microsoft-com:office:office"
      xmlns:w="urn:schemas-microsoft-com:office:word"
      xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"
      xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta charset="utf-8">
<meta name="ProgId" content="Word.Document">
<meta name="Generator" content="Microsoft Word 15">
<title>MD2MSWord Document</title>
<style>
  /* Microsoft Word Document Styles */
  body, p, li, td, th, div {
      font-family: '${targetFont}', sans-serif;
      font-size: 11pt;
      line-height: 1.25;
      color: #000000;
      text-align: left;
      word-spacing: normal;
      letter-spacing: normal;
  }
  p {
      margin: 0 0 6pt 0;
      text-align: left;
  }
  h1 {
      font-family: '${targetFont}', sans-serif;
      font-size: 16pt;
      font-weight: bold;
      color: ${headingColor};
      margin: 14pt 0 6pt 0;
  }
  h2 {
      font-family: '${targetFont}', sans-serif;
      font-size: 13pt;
      font-weight: bold;
      color: ${headingColor};
      margin: 12pt 0 4pt 0;
  }
  h3 {
      font-family: '${targetFont}', sans-serif;
      font-size: 12pt;
      font-weight: bold;
      color: ${headingColor};
      margin: 8pt 0 2pt 0;
  }
  strong, b {
      font-weight: bold;
  }
  em, i {
      font-style: italic;
  }
  code {
      font-family: 'Consolas', monospace;
      font-size: 9.5pt;
      background-color: #F2F2F2;
      padding: 1pt 3pt;
  }
  pre {
      font-family: 'Consolas', monospace;
      font-size: 9.5pt;
      background-color: #F2F2F2;
      padding: 6pt;
      border: 1pt solid #D9D9D9;
      margin: 6pt 0;
  }
  table {
      border-collapse: collapse;
      margin: 8pt 0;
      width: 100%;
  }
  th, td {
      border: 1pt solid #BFBFBF;
      padding: 4pt 6pt;
      text-align: left;
      font-family: '${targetFont}', sans-serif;
      font-size: 11pt;
  }
  th {
      background-color: #F2F2F2;
      font-weight: bold;
      color: ${headingColor};
  }
  blockquote {
      border-left: 3pt solid ${headingColor};
      padding-left: 8pt;
      margin: 6pt 0;
      color: #595959;
      font-style: italic;
  }
  ul, ol {
      margin: 0 0 6pt 0;
      padding-left: 20pt;
  }
  li {
      margin: 2pt 0;
  }
</style>
</head>
<body>
<!--StartFragment-->
${html}
<!--EndFragment-->
</body>
</html>`;

        return wordOfficeHtml;
    }

    // =========================================================================
    // Clipboard Copy Handler
    // =========================================================================

    async function copyForWord() {
        const raw = markdownInput.value;
        if (!raw.trim()) {
            showToast('⚠️ Input is empty. Paste some markdown first.', 'error');
            return;
        }

        const htmlPayload = buildWordHtmlPayload(raw);
        const plainTextPayload = raw;

        let success = false;

        // Try Modern Async Clipboard API (ClipboardItem)
        if (navigator.clipboard && window.ClipboardItem) {
            try {
                const htmlBlob = new Blob([htmlPayload], { type: 'text/html' });
                const textBlob = new Blob([plainTextPayload], { type: 'text/plain' });
                const item = new ClipboardItem({
                    'text/html': htmlBlob,
                    'text/plain': textBlob
                });
                await navigator.clipboard.write([item]);
                success = true;
            } catch (err) {
                console.warn('Async Clipboard API failed, attempting execCommand fallback:', err);
            }
        }

        // Fallback using execCommand copy event
        if (!success) {
            const handleCopy = (e) => {
                e.preventDefault();
                e.clipboardData.setData('text/plain', plainTextPayload);
                e.clipboardData.setData('text/html', htmlPayload);
            };

            document.addEventListener('copy', handleCopy);
            try {
                success = document.execCommand('copy');
            } catch (err) {
                console.error('execCommand copy failed:', err);
            }
            document.removeEventListener('copy', handleCopy);
        }

        if (success) {
            showToast('✅ Copied! Press Ctrl+V in MS Word Desktop to paste native equations.', 'success');
            copyWordBtn.classList.add('copied');
            const originalText = copyWordBtn.querySelector('span').textContent;
            copyWordBtn.querySelector('span').textContent = 'Copied for MS Word!';
            setTimeout(() => {
                copyWordBtn.classList.remove('copied');
                copyWordBtn.querySelector('span').textContent = originalText;
            }, 2000);
        } else {
            showToast('❌ Copy failed. Please check browser clipboard permissions.', 'error');
        }
    }

    // =========================================================================
    // Toast Notification System
    // =========================================================================

    let toastTimer = null;
    function showToast(message, type = 'success') {
        toastNotification.textContent = message;
        toastNotification.className = `toast ${type} show`;
        clearTimeout(toastTimer);
        toastTimer = setTimeout(() => {
            toastNotification.classList.remove('show');
        }, 4000);
    }

    // =========================================================================
    // Helper Functions
    // =========================================================================

    function escapeHtml(str) {
        return str
            .replace(/&/g, '&amp;')
            .replace(/</g, '&lt;')
            .replace(/>/g, '&gt;')
            .replace(/"/g, '&quot;');
    }

    function debounce(fn, delay) {
        let timer;
        return function (...args) {
            clearTimeout(timer);
            timer = setTimeout(() => fn.apply(this, args), delay);
        };
    }

    // =========================================================================
    // Event Listeners
    // =========================================================================

    // Input changes trigger web preview update
    markdownInput.addEventListener('input', debounce(renderWebPreview, 100));

    // Font & Heading Color dropdown changes trigger preview update
    fontSelect.addEventListener('change', renderWebPreview);
    headingColorSelect.addEventListener('change', renderWebPreview);

    // Primary Copy Button
    copyWordBtn.addEventListener('click', copyForWord);

    // Clear Button
    clearBtn.addEventListener('click', () => {
        markdownInput.value = '';
        renderWebPreview();
        markdownInput.focus();
    });

    // Keyboard Shortcut (Ctrl+Enter / Cmd+Enter)
    document.addEventListener('keydown', (e) => {
        if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') {
            e.preventDefault();
            copyForWord();
        }
    });

    // Initialize Preview on Startup
    renderWebPreview();

})();
