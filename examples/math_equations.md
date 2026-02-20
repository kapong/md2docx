---
title: "Math Equations Demo / ตัวอย่างสมการคณิตศาสตร์"
language: th
---

# Math Equations / สมการคณิตศาสตร์

This document demonstrates md2docx support for LaTeX math equations.

เอกสารนี้แสดงตัวอย่างการรองรับสมการคณิตศาสตร์ LaTeX ของ md2docx

## Inline Math / สมการแบบอินไลน์

Einstein's famous equation $E = mc^2$ changed physics forever. The Pythagorean theorem states that $a^2 + b^2 = c^2$ for right triangles. Euler's identity $e^{i\pi} + 1 = 0$ connects five fundamental constants.

The quadratic formula gives us $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ as the solution.

A circle with radius $r$ has area $A = \pi r^2$ and circumference $C = 2\pi r$.

## Display Math / สมการแบบบล็อก

### Basic Equations / สมการพื้นฐาน

$$E = mc^2 \label{eq:einstein}$$

$$a^2 + b^2 = c^2 \label{eq:pythagoras}$$

As shown in Eq. {ref:eq:einstein}, mass and energy are equivalent. The Pythagorean theorem (Eq. {ref:eq:pythagoras}) is fundamental to Euclidean geometry.

### Fractions / เศษส่วน

$$\frac{n!}{k!(n-k)!} = \binom{n}{k}$$

$$\frac{1}{1 + \frac{1}{1 + \frac{1}{x}}}$$

### Greek Letters / อักษรกรีก

$$\alpha + \beta + \gamma = \pi$$

$$\Phi = \frac{1 + \sqrt{5}}{2} \approx 1.618$$

### Summation and Series / ผลรวมและอนุกรม

$$\sum_{i=1}^{n} i = \frac{n(n+1)}{2}$$

$$\sum_{n=0}^{\infty} \frac{x^n}{n!} = e^x$$

$$\prod_{i=1}^{n} i = n!$$

### Integrals / ปริพันธ์

$$\int_0^1 x^2 \, dx = \frac{1}{3}$$

$$\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi} \label{eq:gaussian}$$

$$\int_a^b f(x) \, dx = F(b) - F(a) \label{eq:ftc}$$

The Gaussian integral (Eq. {ref:eq:gaussian}) appears throughout probability theory. The Fundamental Theorem of Calculus (Eq. {ref:eq:ftc}) connects differentiation and integration.

### Square Roots / รากที่สอง

$$x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$$

$$\sqrt[3]{27} = 3$$

### Subscripts and Superscripts / ตัวห้อยและยกกำลัง

$$x_1, x_2, \ldots, x_n$$

$$a_{n+1} = a_n + d$$

$$e^{i\theta} = \cos\theta + i\sin\theta$$

### Matrices / เมทริกซ์

$$\begin{pmatrix} a & b \\ c & d \end{pmatrix} \begin{pmatrix} x \\ y \end{pmatrix} = \begin{pmatrix} ax + by \\ cx + dy \end{pmatrix}$$

$$\begin{bmatrix} 1 & 0 & 0 \\ 0 & 1 & 0 \\ 0 & 0 & 1 \end{bmatrix}$$

### Functions / ฟังก์ชัน

$$\sin^2\theta + \cos^2\theta = 1$$

$$\lim_{x \to 0} \frac{\sin x}{x} = 1$$

$$\log_2 n = \frac{\ln n}{\ln 2}$$

### Accents / เครื่องหมายกำกับ

$$\hat{x}, \bar{y}, \vec{v}, \dot{a}, \ddot{b}, \tilde{n}$$

### Delimiters / วงเล็บ

$$\left( \frac{a}{b} \right) \left[ \frac{c}{d} \right] \left\{ \frac{e}{f} \right\}$$

## Practical Examples / ตัวอย่างจริง

### Maxwell's Equations / สมการแมกซ์เวลล์

$$\nabla \cdot \vec{E} = \frac{\rho}{\epsilon_0}$$

$$\nabla \cdot \vec{B} = 0$$

$$\nabla \times \vec{E} = -\frac{\partial \vec{B}}{\partial t}$$

### Schrödinger Equation / สมการชเรอดิงเงอร์

$$i\hbar\frac{\partial}{\partial t}\Psi = \hat{H}\Psi$$

### Normal Distribution / การแจกแจงปกติ

$$f(x) = \frac{1}{\sigma\sqrt{2\pi}} e^{-\frac{1}{2}\left(\frac{x-\mu}{\sigma}\right)^2} \label{eq:normal}$$

The normal distribution (Eq. {ref:eq:normal}) uses the Gaussian integral from Eq. {ref:eq:gaussian}.

### Cauchy-Schwarz Inequality / อสมการโคชี-ชวาร์ตซ์

$$\left( \sum_{k=1}^n a_k b_k \right)^2 \leq \left( \sum_{k=1}^n a_k^2 \right) \left( \sum_{k=1}^n b_k^2 \right)$$

## Advanced Constructs / โครงสร้างขั้นสูง

### Piecewise Functions / ฟังก์ชันแบบแยกส่วน

$$f(x) = \left\{ \begin{array}{ll} x^2 & \text{if } x \geq 0 \\ -x & \text{if } x < 0 \end{array} \right.$$

$$|x| = \left\{ \begin{array}{rl} x & x \geq 0 \\ -x & x < 0 \end{array} \right.$$

$$\text{sgn}(x) = \left\{ \begin{array}{rl} 1 & x > 0 \\ 0 & x = 0 \\ -1 & x < 0 \end{array} \right.$$

### Aligned Equations / สมการจัดเรียง

$$\begin{aligned} f(x) &= x^2 + 2x + 1 \\ &= (x+1)^2 \end{aligned}$$

$$\begin{aligned} \nabla \cdot \vec{E} &= \frac{\rho}{\epsilon_0} \\ \nabla \cdot \vec{B} &= 0 \\ \nabla \times \vec{E} &= -\frac{\partial \vec{B}}{\partial t} \\ \nabla \times \vec{B} &= \mu_0 \vec{J} + \mu_0 \epsilon_0 \frac{\partial \vec{E}}{\partial t} \end{aligned}$$

### Nth Roots / รากที่ n

$$\sqrt[3]{27} = 3, \quad \sqrt[4]{256} = 4, \quad \sqrt[5]{32} = 2$$

$$\sqrt[n]{a \cdot b} = \sqrt[n]{a} \cdot \sqrt[n]{b}$$

$$\sqrt[3]{\sqrt{x^2 + 1}}$$

### Overbrace and Underbrace / วงเล็บบนและล่าง

$$\overbrace{a + a + \cdots + a}^{n}$$

$$\underbrace{1 + 1 + \cdots + 1}_{n} = n$$

### Determinants / ดีเทอร์มิแนนต์

$$\begin{vmatrix} a & b \\ c & d \end{vmatrix} = ad - bc$$

$$\begin{Vmatrix} a & b \\ c & d \end{Vmatrix}$$

$$\begin{vmatrix} a_{11} & a_{12} & a_{13} \\ a_{21} & a_{22} & a_{23} \\ a_{31} & a_{32} & a_{33} \end{vmatrix}$$

### Matrix Types / ชนิดเมทริกซ์

$$\begin{bmatrix} 1 & 2 \\ 3 & 4 \end{bmatrix} \quad \begin{Bmatrix} a & b \\ c & d \end{Bmatrix} \quad \begin{vmatrix} e & f \\ g & h \end{vmatrix}$$

### Substack / ผลรวมหลายเงื่อนไข

$$\sum_{\substack{0 < i < m \\ 0 < j < n}} P(i,j)$$

### Floor and Ceiling / ฟังก์ชันปัดลงและปัดขึ้น

$$\lfloor x \rfloor + \lceil y \rceil$$

$$\lfloor \pi \rfloor = 3, \quad \lceil e \rceil = 3$$

### Deep Nested Fractions / เศษส่วนซ้อนลึก

$$\frac{1}{1 + \frac{1}{1 + \frac{1}{1 + \frac{1}{x}}}}$$

### Tensor and Multi-Index / เทนเซอร์และดัชนีซ้อน

$$R^{\mu}{}_{\nu\rho\sigma}$$

$$T^{\alpha\beta} = g^{\alpha\mu} g^{\beta\nu} T_{\mu\nu}$$

## Physics and Engineering / ฟิสิกส์และวิศวกรรม

### Dirac Notation / สัญกรณ์ดิแรก

$$\langle \psi | \hat{H} | \phi \rangle$$

$$|\psi\rangle = \alpha|0\rangle + \beta|1\rangle$$

### Fourier Transform / การแปลงฟูเรียร์

$$\hat{f}(\xi) = \int_{-\infty}^{\infty} f(x) e^{-2\pi i x \xi} dx$$

### Laplacian / ลาปลาเซียน

$$\nabla^2 f = \frac{\partial^2 f}{\partial x^2} + \frac{\partial^2 f}{\partial y^2}$$

### Maxwell's Equations (curl form)

$$\nabla \times \vec{B} = \mu_0 \vec{J} + \mu_0 \epsilon_0 \frac{\partial \vec{E}}{\partial t}$$

### Stirling's Approximation / สูตรประมาณสเตอร์ลิง

$$n! \approx \sqrt{2\pi n} \left(\frac{n}{e}\right)^n$$

### Taylor Series / อนุกรมเทย์เลอร์

$$f(x) = \sum_{n=0}^{\infty} \frac{f^{(n)}(a)}{n!}(x-a)^n$$

### Bayes' Theorem / ทฤษฎีบทเบย์

$$P(A|B) = \frac{P(B|A) \cdot P(A)}{P(B)} \label{eq:bayes}$$

Bayes' theorem (Eq. {ref:eq:bayes}) is central to statistical inference and works alongside the normal distribution (Eq. {ref:eq:normal}).

### Euler Product Formula / สูตรผลคูณออยเลอร์

$$\zeta(s) = \sum_{n=1}^{\infty} \frac{1}{n^s} = \prod_{p} \frac{1}{1 - p^{-s}}$$

### Residue Theorem / ทฤษฎีบทเรซิดิว

$$\oint_{\gamma} f(z) \, dz = 2\pi i \sum_{k=1}^{n} \text{Res}(f, a_k)$$

### Binomial Coefficient / สัมประสิทธิ์ทวินาม

$$\binom{n}{k} = \frac{n!}{k!(n-k)!}$$

$$(x+y)^n = \sum_{k=0}^{n} \binom{n}{k} x^{n-k} y^k$$

### Font Variants / รูปแบบตัวอักษร

$$\mathbb{R}^n, \quad \mathbf{v} = (v_1, v_2, v_3), \quad \mathcal{L}\{f(t)\} = F(s)$$
