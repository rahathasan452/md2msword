# 📝 MD2MSWord — AI Markdown & Math to MS Word Desktop Engine

[![Live Website](https://img.shields.io/badge/Live_App-GitHub_Pages-2EA44F?style=for-the-badge&logo=github)](https://rahathasan452.github.io/md2msword/)
[![MS Word Desktop](https://img.shields.io/badge/Target-MS_Word_Desktop-0078D4?style=for-the-badge&logo=microsoftword)](https://www.microsoft.com/en-us/microsoft-365/word)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg?style=for-the-badge)](https://opensource.org/licenses/MIT)

**MD2MSWord** is a high-performance, 100% client-side web converter designed to take Markdown responses and LaTeX math equations from AI models (**ChatGPT**, **Claude**, **Gemini**, **DeepSeek**) and copy them directly into **Microsoft Word Desktop software** as 100% native, editable Word Equations (`m:oMath`) and formatted Office elements.

🌐 **Live Web Application**: [https://rahathasan452.github.io/md2msword/](https://rahathasan452.github.io/md2msword/)

---

> [!IMPORTANT]
> ### ⚠️ CRITICAL COMPATIBILITY DISCLAIMER
> **This tool is specifically engineered for installed Microsoft Word Desktop software** (Microsoft Word 365, Word 2024, Word 2021, Word 2019, Word 2016 for Windows & macOS).
> 
> ❌ **It will NOT work in web apps like Word Online (Word Web App) or Google Docs**, because web browser editors strip system clipboard MathML equation tags and Office table namespaces upon pasting.

---

## ⚡ The Core Problem Solved

When copying raw AI output containing KaTeX HTML into MS Word Desktop:
1. **Character Doubling**: Standard KaTeX renders both visual HTML (`<span class="katex-html">`) and MathML (`<math>`). Pasting this into Word causes duplicated text (e.g. `I_A R_A I_A R_A`).
2. **Missing Equation Editor**: Plain Markdown text pastes as unformatted plain text without activating Word's native equation builder (`m:oMath`).
3. **Broken Blockquotes & Checkboxes**: Standard Markdown blockquotes paste as unindented text, and task list checkboxes paste as plain bullet points.

### How MD2MSWord Fixes This:
- **MathML Extraction**: Extracts pure MathML `<math xmlns="http://www.w3.org/1998/Math/MathML">` and strips all visual KaTeX HTML nodes.
- **Word HTML Wrapping**: Wraps the output with Microsoft Office XML namespaces (`xmlns:w`, `xmlns:m`, `xmlns:o`) and Windows clipboard fragment markers (`<!--StartFragment-->`).
- **Office 1-Cell Table Converter**: Converts GitHub callout alerts and nested blockquotes into 1-cell Office HTML tables, forcing Word Desktop to render **native vertical left accent bars** and **shaded background cards**.

---

## 🚀 Key Features & Capabilities

- ⚡ **Native Editable Word Equations**: Pasted equations automatically convert into native, fully-editable Word Equations (`m:oMath`).
- 🤖 **AI Math Delimiter Protection**: Handles unescaped bracket `[` / `]` and parenthesis `(` / `)` math formats copied directly from ChatGPT, Gemini, Claude, and DeepSeek.
- 🎨 **Office Callout Alert Boxes**: Converts `> [!NOTE]`, `> [!TIP]`, `> [!IMPORTANT]`, `> [!WARNING]`, and `> [!CAUTION]` into styled callout alert boxes with left accent borders and soft background tints.
- 🧱 **Depth-Aware Nested Blockquotes**: Multi-level blockquotes (`> Level 1` $\to$ `> > Level 2` $\to$ `> > > Level 3`) paste with distinct left vertical accent bars and shaded card backgrounds.
- ☑️ **Native Word Checkbox Symbols**: Converts GFM task lists (`- [x]` and `- [ ]`) into green checked ballot boxes (`☑`) and open ballot boxes (`☐`).
- ⬚ **Zero Dotted-Box Placeholder Bug Fix**: Injects zero-width spaces (`&#x200B;`) into empty MathML base elements to prevent MS Word from rendering dotted square placeholders (⬚) on nuclear isotopes and presubscripts.
- 🔤 **Custom Typography & Colors**: Choose from 8 target fonts (*Calibri*, *Aptos*, *Times New Roman*, *Georgia*, *Segoe UI*, *Arial*, *Cambria*, *Verdana*) and 7 heading accent colors (*Word Navy*, *Crimson Academic*, *Deep Teal*, etc.).
- 🧪 **15-Part Stress Test Suite**: Comes with a built-in benchmark file (`stress_test.md`) covering matrices, vector calculus, quantum decay, probability, and complex GFM tables.

---

## 🛠️ How to Use

1. Open the live web app: **[https://rahathasan452.github.io/md2msword/](https://rahathasan452.github.io/md2msword/)**
2. Paste Markdown text from ChatGPT, Claude, Gemini, or DeepSeek into the left editor panel.
3. Choose your target **Word Font** and **Heading Accent Color** from the top control strip.
4. Click **Copy for MS Word** (or press <kbd>Ctrl</kbd> + <kbd>Enter</kbd>).
5. Open **installed Microsoft Word Desktop software**, press <kbd>Ctrl</kbd> + <kbd>V</kbd> to paste, and ensure **Keep Source Formatting (K)** is selected in Word's Paste Options popup menu if prompted.

---

## 💻 Local Setup & Deployment

### Run Locally:
```bash
# Clone the repository
git clone https://github.com/rahathasan452/md2msword.git

# Navigate into the project folder
cd md2msword

# Open index.html in any web browser
```

### Deploy to GitHub Pages:
1. Create a public repository named `md2msword` on GitHub.
2. Push your code: `git push -u origin main`.
3. Go to **Settings** $\rightarrow$ **Pages** $\rightarrow$ Select **`main`** branch $\rightarrow$ Click **Save**.

---

## 🔬 Tech Stack

- **GFM Parser**: [marked.js](https://marked.js.org/)
- **TeX Math Engine**: [KaTeX](https://katex.org/)
- **Frontend Architecture**: Vanilla HTML5, CSS3, JavaScript (Zero build step, 100% client-side)
- **Hosting**: GitHub Pages

---

## 📄 License

This project is licensed under the [MIT License](LICENSE).
