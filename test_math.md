# ทดสอบสมการคณิตศาสตร์

2. **การวิเคราะห์สัญญาณการเคลื่อนไหวและความสอดคล้อง:** ระบบแปลงเวกเตอร์ท่าทางให้กลายเป็นกราฟ 1 มิติ 2 เส้น ได้แก่ Activity Signal ที่คำนวณจากขนาดของความเร็วในปริภูมิแฝง (Rabiner & Schafer, 1978):

   $$A(t) = \|\mathbf{z}_t - \mathbf{z}_{t-1}\|_2$$

   และ Coherence Signal ที่วัดค่าเฉลี่ยของ Cosine Similarity ระหว่างเฟรม $t$ กับเฟรมข้างเคียงภายในรัศมี $W$ (Shechtman & Irani, 2007):

   $$C(t) = \frac{1}{|\mathcal{N}(t)|} \sum_{\tau \in \mathcal{N}(t)} \text{CosSim}(\hat{\mathbf{z}}_t,\; \hat{\mathbf{z}}_\tau)$$

   โดย $\mathcal{N}(t) = \{\tau : |\tau - t| \le W\}$ คือเซตเพื่อนบ้าน และ $\hat{\mathbf{z}}_t = \mathbf{z}_t / \|\mathbf{z}_t\|$ คือเวกเตอร์ที่ทำให้เป็นหนึ่งหน่วย (Unit-normalized)
