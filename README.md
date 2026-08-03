# 📝 MD2MSWord — AI Markdown & Math to MS Word Desktop Converter

[![Live Website](https://img.shields.io/badge/Live_App-GitHub_Pages-2EA44F?style=for-the-badge&logo=github)](https://rahathasan452.github.io/md2msword/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

**MD2MSWord** is a high-performance client-side web application designed to copy AI responses (**ChatGPT**, **Claude**, **Gemini**, **DeepSeek**) containing Markdown and LaTeX math equations directly into **Microsoft Word Desktop software** with 100% native, editable MS Word Math Equations (`m:oMath`).

🌐 **Live Web Application**: [https://rahathasan452.github.io/md2msword/](https://rahathasan452.github.io/md2msword/)

---

> [!IMPORTANT]
> ### ⚠️ Compatibility Disclaimer
> **This tool ONLY works with installed Microsoft Word Desktop software** (Microsoft Word 365, Word 2024, Word 2021, Word 2019, Word 2016 on Windows & macOS).
> 
> ❌ **It will NOT work in web apps like Word Online (Word Web App) or Google Docs**, as cloud browser editors strip clipboard MathML equations and Office table namespaces upon pasting.

---

## ⚡ The Problem Solved

When copying AI output containing KaTeX HTML into MS Word Desktop:
- Default KaTeX visual nodes (`<span class="katex-html">`) combined with MathML (`<math>`) cause text duplication in Word (e.g., `I_A R_A I_A R_A`).
- **MD2MSWord** extracts MathML `<math xmlns="http://www.w3.org/1998/Math/MathML">` and strips redundant KaTeX visual nodes before building the system clipboard payload.
- When pasted, **MS Word Desktop** converts every `<math>` block into a native, fully-editable Word Equation (`m:oMath`) without any character duplication.

---

## 🚀 Key Features

- **100% Client-Side Engine**: Runs directly in the browser with zero backend dependencies or server setup required.
- **AI Math Delimiter Protection**: Supports unescaped bracket `[` / `]` and parenthesis `(` / `)` math delimiters from ChatGPT, Gemini, Claude, and DeepSeek.
- **Office 1-Cell Callout Tables**: Converts GitHub callout alerts (`> [!NOTE]`, `> [!WARNING]`, etc.) and nested blockquotes (`> Level 1` $\to$ `> > Level 2`) into native 1-cell Word tables with left vertical accent bars, proper step indentations, and background shading cards.
- **Word Checkbox Symbols**: Converts GFM task lists (`- [x]` and `- [ ]`) into native Word checked (`☑`) and unchecked (`☐`) ballot symbols.
- **Zero Dotted-Box Placeholder Bug**: Automatically injects zero-width spaces (`&#x200B;`) into empty MathML base elements to prevent MS Word from rendering dotted square placeholders (⬚).
- **Customizable Export Styling**: Custom target fonts (*Calibri*, *Aptos*, *Times New Roman*, *Georgia*, *Segoe UI*, *Arial*, *Cambria*, *Verdana*) and heading accent colors with live color swatch previews.
- **Built-in 15-Part Stress Test Suite**: Includes a comprehensive stress test document (`stress_test.md`) covering matrices, chemical decay, vector calculus, probability, and GFM elements.

---

## 🛠️ How to Use

1. Open the live web app: [https://rahathasan452.github.io/md2msword/](https://rahathasan452.github.io/md2msword/)
2. Paste Markdown text from ChatGPT / Claude / DeepSeek in the left editor panel.
3. Click **Copy for MS Word** (or press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>).
4. Press <kbd>Ctrl</kbd> + <kbd>V</kbd> inside **installed Microsoft Word Desktop software**.

---

## 💻 Local Development

```bash
# Clone the repository
git clone https://github.com/rahathasan452/md2msword.git

# Navigate to directory
cd md2msword

# Open index.html in your browser
```

---

### Tech Stack
- [marked.js](https://marked.js.org/) — GFM Markdown Parser
- [KaTeX](https://katex.org/) — TeX to MathML conversion engine
- Vanilla HTML5 / CSS3 / JavaScript — Zero build tools required
- GitHub Pages — Free static site hosting

---

### License
This project is open source and available under the [MIT License](LICENSE).
