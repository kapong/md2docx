---
title: "Math Equations Demo / ตัวอย่างสมการคณิตศาสตร์"
language: en
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

$$E = mc^2$$

$$a^2 + b^2 = c^2$$

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

$$\int_{-\infty}^{\infty} e^{-x^2} dx = \sqrt{\pi}$$

$$\int_a^b f(x) \, dx = F(b) - F(a)$$

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

$$f(x) = \frac{1}{\sigma\sqrt{2\pi}} e^{-\frac{1}{2}\left(\frac{x-\mu}{\sigma}\right)^2}$$

### Cauchy-Schwarz Inequality / อสมการโคชี-ชวาร์ตซ์

$$\left( \sum_{k=1}^n a_k b_k \right)^2 \leq \left( \sum_{k=1}^n a_k^2 \right) \left( \sum_{k=1}^n b_k^2 \right)$$
