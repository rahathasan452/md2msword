# 🧪 MD2MSWord Ultimate Markdown Stress Test Suite

Use this comprehensive 11-part test suite to stress test **MD2MSWord** across every feature that can exist in Markdown files, GFM extensions, LaTeX math, HTML elements, and AI copy responses.

---

## 🔹 Test Suite 1: ChatGPT & AI Copy Delimiters

Copy and paste this section to verify that ChatGPT's unescaped bracket `[` / `]` and parenthesis `(` / `)` math delimiters convert properly:

### 1.1 ChatGPT Display Brackets:
[
\left|\frac{v_o}{v_{in}}\right| = \frac{A_F}{\sqrt{1+\left(\frac{f}{f_H}\right)^2}}
]

[
\phi = -\tan^{-1}\left(\frac{f}{f_H}\right)
]

### 1.2 ChatGPT Inline Parentheses:
- Resistance relation: ( a = 1 ) and ( b = \frac{f}{f_H} )
- High frequency condition: ( f \gg f_H ) leads to phase shift ( \phi \to -90^\circ )
- Transfer function: ( H(j\omega) = \frac{1}{1 + jx} ) where ( x = \frac{f}{f_H} )
- Angle definition: ( \theta = \tan^{-1}(x) )

---

## 🔹 Test Suite 2: Advanced LaTeX Math, Matrices & Accents

### 2.1 Aligned Multi-line System of Equations:
$$
\begin{aligned}
\nabla \times \mathbf{E} &= -\frac{\partial \mathbf{B}}{\partial t} \\
\nabla \times \mathbf{B} &= \mu_0 \mathbf{J} + \mu_0 \varepsilon_0 \frac{\partial \mathbf{E}}{\partial t} \\
\nabla \cdot \mathbf{E} &= \frac{\rho}{\varepsilon_0} \\
\nabla \cdot \mathbf{B} &= 0
\end{aligned}
$$

### 2.2 Piecewise Functions (Cases):
$$
f(x) = \begin{cases} 
0, & \text{if } x < 0 \\
\frac{1}{2}x^2, & \text{if } 0 \le x \le 1 \\
1 - e^{-x}, & \text{if } x > 1 
\end{cases}
$$

### 2.3 Multiple Matrix Types:
- Parentheses Matrix ($2 \times 2$):
  $$\mathbf{P} = \begin{pmatrix} a & b \\ c & d \end{pmatrix}$$
- Bracketed Matrix ($3 \times 3$):
  $$\mathbf{B} = \begin{bmatrix} 1 & 0 & 0 \\ 0 & 1 & 0 \\ 0 & 0 & 1 \end{bmatrix}$$
- Determinant Bar Matrix:
  $$\det(\mathbf{M}) = \begin{vmatrix} x_1 & y_1 \\ x_2 & y_2 \end{vmatrix} = x_1 y_2 - x_2 y_1$$

### 2.4 Binomial Coefficients & Vector Accents:
$$
\binom{n}{k} = \frac{n!}{k!(n-k)!}, \quad \vec{F} = m \vec{a} = m \frac{d^2 \mathbf{r}}{dt^2}, \quad \hat{y} = \bar{x} + \tilde{\varepsilon}
$$

---

## 🔹 Test Suite 3: GFM Task Lists & Checkboxes

- [x] Implement LaTeX math pre-parser engine
- [x] Strip KaTeX `.katex-html` duplicate visual nodes
- [x] Generate Word Office MathML (`<math xmlns="...">`)
- [ ] Add export to PDF feature
- [ ] Add direct cloud sync

---

## 🔹 Test Suite 4: Rich Typography & Text Styling

- **Bold Text**: **Strong emphasis with double asterisks** or __double underscores__
- *Italic Text*: *Single asterisk emphasis* or _single underscore emphasis_
- ***Bold & Italic***: ***Combined emphasis text***
- ~~Strikethrough~~: ~~This feature is deprecated~~
- <u>Underlined Text</u>: <u>Underlined for emphasis</u>
- <mark>Highlighted Text</mark>: <mark>Important highlighted note</mark>
- Subscript & Superscript: H<sub>2</sub>O, CO<sub>2</sub>, $x^2$, $e^{i\pi} + 1 = 0$
- Keyboard Shortcut Tags: Press <kbd>Ctrl</kbd> + <kbd>Enter</kbd> to copy for MS Word

---

## 🔹 Test Suite 5: GitHub Callout Alerts & Blockquotes

> [!NOTE]
> This is a helpful background note regarding Microsoft Word MathML compatibility.

> [!TIP]
> Use shortcut <kbd>Ctrl</kbd>+<kbd>Enter</kbd> to copy instantly to your clipboard.

> [!IMPORTANT]
> Always use MS Word Desktop software for 100% native equation rendering.

> [!WARNING]
> Web apps like Google Docs or Word Online do not parse clipboard MathML.

> [!CAUTION]
> Avoid manually editing XML tags inside generated Office document clipboard structures.

### Nested Blockquotes:
> Level 1: Outer blockquote description for system architecture.
> > Level 2: Nested blockquote containing math equation:
> > $$E = mc^2 = \frac{m_0 c^2}{\sqrt{1 - v^2/c^2}}$$
> > > Level 3: Deeply nested note regarding relativistic momentum.

---

## 🔹 Test Suite 6: Collapsible Accordions & HTML Elements

<details>
<summary><b>Click to expand Advanced Derivation Notes</b></summary>

Here is the expanded technical proof:

$$
\int_{0}^{\infty} x^{n} e^{-ax} \, dx = \frac{\Gamma(n+1)}{a^{n+1}} = \frac{n!}{a^{n+1}}
$$

- Parameter $a > 0$
- Integer $n \ge 0$

</details>

---

## 🔹 Test Suite 7: HTML Tables with Colspan & Rowspan

<table border="1" cellpadding="6" cellspacing="0">
  <thead>
    <tr>
      <th colspan="2">System Specifications</th>
      <th>Measured Output</th>
      <th>Status</th>
    </tr>
  </thead>
  <tbody>
    <tr>
      <td rowspan="2"><b>Power Unit</b></td>
      <td>Voltage Supply</td>
      <td>$230\text{ V} \pm 5\%$</td>
      <td>✅ Normal</td>
    </tr>
    <tr>
      <td>Current Draw</td>
      <td>$45.2\text{ A}$</td>
      <td>⚡ Active</td>
    </tr>
    <tr>
      <td colspan="2"><b>Overall Efficiency</b></td>
      <td>$\eta = 94.6\%$</td>
      <td>🏆 High Performance</td>
    </tr>
  </tbody>
</table>

---

## 🔹 Test Suite 8: Code Blocks with Syntax Highlighting

### Python (Scientific Computation):
```python
import numpy as np
import scipy.linalg as la

def solve_eigenvalues(matrix_a):
    """Calculates eigenvalues and eigenvectors for input matrix."""
    eigenvalues, eigenvectors = la.eig(matrix_a)
    return np.real(eigenvalues), eigenvectors
```

### C++ (Low-level Hardware Controller):
```cpp
#include <iostream>
#include <cmath>

struct Vector3D {
    double x, y, z;
    double magnitude() const {
        return std::sqrt(x*x + y*y + z*z);
    }
};

int main() {
    Vector3D force{3.0, 4.0, 12.0};
    std::cout << "Force Magnitude: " << force.magnitude() << " N\n";
    return 0;
}
```

### SQL Query:
```sql
SELECT 
    device_id,
    AVG(voltage_reading) AS avg_voltage,
    MAX(current_peak) AS max_current
FROM sensor_logs
WHERE timestamp >= NOW() - INTERVAL '24 hours'
GROUP BY device_id
HAVING AVG(voltage_reading) > 220.0;
```

---

## 🔹 Test Suite 9: Footnotes, Links & Images

### Links & Autolinks:
- Direct Link: [Microsoft Word Official Site](https://www.microsoft.com/en-us/microsoft-365/word)
- Autolink: <https://katex.org/>
- Relative Reference Link: [MD2MSWord Documentation][ref1]

[ref1]: file:///d:/Project/md2msword/README.md "MD2MSWord Readme"

---

## 🔹 Test Suite 10: Chemical Equations & Physics Reactions

### Chemistry Equations:
$$
\text{2 H}_2\text{ (g)} + \text{O}_2\text{ (g)} \longrightarrow \text{2 H}_2\text{O (l)}, \quad \Delta H = -571.6\text{ kJ/mol}
$$

$$
\text{CaCO}_3\text{ (s)} + \text{2 HCl (aq)} \longrightarrow \text{CaCl}_2\text{ (aq)} + \text{CO}_2\text{ (g)} + \text{H}_2\text{O (l)}
$$

### Nuclear Decay:
$$
{}_{92}^{238}\text{U} \longrightarrow {}_{90}^{234}\text{Th} + {}_{2}^{4}\alpha + \gamma
$$

---

## 🔹 Test Suite 11: Real-World Master Benchmark Document

# 1. Potential Transformer (PT) & Current Transformer (CT) Analysis

Instrument transformers provide galvanically isolated interfaces between high-power transmission lines and low-voltage measuring instruments.

## 1.1 Potential Transformer (PT)

A Potential Transformer is connected in parallel (shunt) across the high-voltage supply. It operates as a precision step-down transformer, reducing high line voltages down to standardized lower voltage levels (typically $110\text{ V}$ or lower) suitable for the voltage coil (potential coil) of a standard wattmeter.

## 1.2 Current Transformer (CT)

A Current Transformer is connected in series with the load phase line. It features a primary winding with very few turns of heavy cross-section conductor carrying full line current, and a secondary winding with a larger number of turns. It steps down large line currents to standardized secondary values (typically $5\text{ A}$ or $1\text{ A}$) to energize the current coil of the wattmeter.

## 1.3 Indirect Active Power Calculation

When using instrument transformers, the actual active power ($P_M$) consumed by the primary circuit is derived by multiplying the power measured on the secondary wattmeter ($P_1$) by both transformation ratios:

$$
P_M = P_1 \times k_{\text{CT}} \times k_{\text{PT}}
$$

Where:
* $P_M$: Total calculated primary active power ($\text{W}$)
* $P_1$: Active power reading indicated by secondary wattmeter ($\text{W}$)
* $k_{\text{CT}}$: Current Transformer transformation ratio ($k_{\text{CT}} = \frac{I_1}{I_2}$)
* $k_{\text{PT}}$: Potential Transformer transformation ratio ($k_{\text{PT}} = \frac{V_1}{V_2}$)

---

## 🔹 Test Suite 12: Trigonometric & Complex Analysis

### 12.1 Euler's Identity & De Moivre's Theorem:
$$
e^{i\pi} + 1 = 0, \quad (\cos \theta + i \sin \theta)^n = \cos(n\theta) + i \sin(n\theta)
$$

### 12.2 Fourier & Laplace Integral Transforms:
$$
\hat{f}(\xi) = \int_{-\infty}^{\infty} f(x) \, e^{-2\pi i x \xi} \, dx
$$

$$
\mathcal{L}\{f(t)\} = F(s) = \int_{0}^{\infty} f(t) \, e^{-st} \, dt
$$

---

## 🔹 Test Suite 13: Linear Algebra & Vector Physics

### 13.1 Vector Products:
$$
\vec{A} \cdot \vec{B} = \|\vec{A}\| \|\vec{B}\| \cos\theta, \quad \vec{A} \times \vec{B} = \mathbf{\hat{n}} \|\vec{A}\| \|\vec{B}\| \sin\theta
$$

### 13.2 Tensor Field Stress-Energy Tensor:
$$
T^{\mu\nu} = \partial^{\mu}\phi \partial^{\nu}\phi - \frac{1}{2} g^{\mu\nu} \left( \partial_{\rho}\phi \partial^{\rho}\phi - m^2 \phi^2 \right)
$$

---

## 🔹 Test Suite 14: Probability & Statistics

### 14.1 Gaussian Normal Distribution:
$$
f(x \mid \mu, \sigma^2) = \frac{1}{\sqrt{2\pi\sigma^2}} \exp\left(-\frac{(x - \mu)^2}{2\sigma^2}\right)
$$

### 14.2 Bayes' Theorem & Expectation:
$$
P(A \mid B) = \frac{P(B \mid A) P(A)}{P(B)}, \quad \mathbb{E}[X] = \int_{-\infty}^{\infty} x f(x) \, dx
$$

---

## 🔹 Test Suite 15: Tables with Embedded Math & Inline Badges

| Component | Function / Formula | Operating Specs | Status |
| :--- | :--- | :---: | :---: |
| **Step-Down PT** | $V_2 = V_1 \left(\frac{N_2}{N_1}\right)$ | $110\text{ V}$ Nominal | <mark>Active</mark> |
| **Series CT** | $I_2 = I_1 \left(\frac{N_1}{N_2}\right)$ | $5\text{ A}$ Nominal | ✅ Calibrated |
| **Nuclear Decay** | ${}_{92}^{238}\text{U} \to {}_{90}^{234}\text{Th} + {}_{2}^{4}\alpha$ | $t_{1/2} = 4.47 \times 10^9\text{ yr}$ | ⚛️ Stable Isotope |

