# 數論（一）期末報告 — Ring-LWE 與理想格密碼學

> Algebraic Number Theory Meets Post-Quantum Security
>
> 黃崇晉 (Chung-Chin Huang) · L16141149 · 數論（一）· 2026 春季 · 國立成功大學

## 檔案結構

| 檔案 | 說明 |
|---|---|
| `presentation_v2.pptx` / `presentation_v2.pdf` | 主報告投影片（41 頁） |
| `presentation_script_v2.tex` / `.pdf` | 報告講稿 |
| `make_ppt_v2.py` | 投影片生成程式（python-pptx） |
| `summary.pdf` | 5 頁書面摘要 |
| `補充說明.pdf` | Period Finding / Integer Factoring / Lattice / SVP / CVP 補充說明 |
| `minkowski_viz.png` | Minkowski 第一定理視覺化圖（n=2） |
| `speaker_script.txt` | 口頭報告草稿 |

## 編譯方式

LaTeX 來源以 `xelatex` 編譯（需 xeCJK 與 Microsoft JhengHei 字型）：
```bash
xelatex summary.tex
xelatex presentation_script_v2.tex
```
投影片重新生成：
```bash
python make_ppt_v2.py
```

---

## Motivation

古典公鑰密碼學（RSA、ECC）將被量子電腦上的 **Shor 演算法**（1994）擊潰，因兩者皆可化為尋找週期（period finding）。RSA 安全性建基於整數分解：取 $a \in \mathbb{Z}_N$ 與 $\gcd(a,N)=1$，由 $f(x)=a^x \bmod N$ 的週期 $\omega$ 得 $N \mid (a^\omega - 1) \Rightarrow pq \mid (a^{\omega/2}+1)(a^{\omega/2}-1)$，再用 gcd 拆出 $p,q$。ECC 之 ECDLP 同樣歸約至離散對數的週期問題。

**古典 vs 量子複雜度：**

| 問題 | 古典最佳 | 量子（Shor） |
|---|---|---|
| 整數分解（RSA-2048） | GNFS $\exp(\tilde{O}((\log N)^{1/3}))$ | $\tilde{O}((\log N)^3)$ |
| ECDLP（256-bit） | Pollard rho $O(\sqrt{n})$ | $\tilde{O}((\log n)^3)$ |
| $\text{SVP}_\gamma$ on lattices | BKZ + Sieving $2^{\Theta(n)}$ | $2^{\Theta(n)}$（未知有效量子演算法） |

**Harvest Now, Decrypt Later**：今日攔截的密文可待量子電腦上線後解密——醫療紀錄、外交電文、長期密鑰若加密生命週期超過量子電腦問世時間等同已暴露。NIST 預估 2030–2035 容錯量子電腦可達破解規模，故 PQC 必須提前部署。**格密碼學**（Lattice-based cryptography）提供後量子安全，其困難性可歸約至格上的最壞情形幾何問題——即便對量子敵手仍認為困難。歷史脈絡：Ajtai（1996）給出第一個 worst-case → average-case 歸約；Regev（2005）提出 LWE，但公鑰大 $O(n^2)$；Lyubashevsky–Peikert–Regev（2010）提出 Ring-LWE，把這份安全性結合代數數論的環結構，公鑰降至 $O(n)$、乘法 $O(n \log n)$（NTT），達到實用效率。

---

## Step 1 — Lattices and Geometry of Numbers

**格的定義**：設 $b_1,\ldots,b_n \in \mathbb{R}^n$ 線性獨立，則

$$\Lambda(B) = \bigoplus_{i=1}^{n} \mathbb{Z}\, b_i = \left\{\sum_{i=1}^n m_i b_i : m_i \in \mathbb{Z}\right\} \subset \mathbb{R}^n$$

為離散加法子群。基本不變量包括**覆積** $\det(\Lambda) = |\det B|$（基本胞體積）與**相繼極小** $\lambda_i(\Lambda)$，其中 $\lambda_1(\Lambda)$ 為最短非零向量長度。

**SVP — Shortest Vector Problem**：求 $x \in \Lambda(B) \setminus \{0\}$ 使 $\|x\| = \lambda_1(\Lambda)$ 最小。直觀為「在離散點集中找離原點最近的非零點」；基底歪扭時即使最短向量很短，亦難從基底辨識，需先做 lattice basis reduction（如 LLL）。

**CVP — Closest Vector Problem**：給目標 $y \in \mathbb{R}^n$，求 $x \in \Lambda(B)$ 使 $\|x-y\|$ 最小。CVP 至少和 SVP 一樣難（SVP 可規約至 CVP）；目前最佳古典與量子演算法在小 $\gamma$ 下需 $2^{\Theta(n)}$ 時間。

**Minkowski 第一定理（1889）**：

> 若 $S \subset \mathbb{R}^n$ 為中心對稱凸體（$x \in S \Rightarrow -x \in S$，任兩點線段含於 $S$）且 $\text{vol}(S) > 2^n \det(\Lambda)$，則 $S$ 必含一非零格點 $v \in \Lambda \cap S \setminus \{0\}$。

**SVP 上界推導**：取中心對稱立方體 $C_r = [-r,r]^n$，$\text{vol}(C_r) = (2r)^n$。當 $r > \det(\Lambda)^{1/n}$ 時由 Minkowski 知 $C_r$ 含非零格點 $v$，故 $\|v\|_\infty \leq r$。配合 $\|v\|_2 \leq \sqrt{n}\,\|v\|_\infty$ 並令 $r \searrow \det(\Lambda)^{1/n}$ 得：

$$\boxed{\lambda_1(\Lambda) \leq \sqrt{n} \cdot \det(\Lambda)^{1/n}} \quad \text{（連結格幾何與課程框架——Obj. 5）}$$

---

## Step 2 — Number Rings as Ideal Lattices

**分圓域**：設 $\zeta_n = e^{2\pi i/n}$ 為本原 $n$ 次單位根，分圓多項式

$$\Phi_n(x) = \prod_{\substack{1\leq k\leq n\\\gcd(k,n)=1}} (x - \zeta_n^k) \in \mathbb{Z}[x]$$

為 monic、$\mathbb{Q}$ 上 irreducible，$\deg \Phi_n = \varphi(n)$。分圓域 $K = \mathbb{Q}(\zeta_n)$ 之 $[K:\mathbb{Q}] = \varphi(n)$，整數環 $\mathcal{O}_K = \mathbb{Z}[\zeta_n]$ 為 **Dedekind domain**（Noetherian、整封閉、Krull 維度 1），每個非零理想唯一分解為質理想之積 $\mathfrak{a} = \mathfrak{p}_1^{e_1} \cdots \mathfrak{p}_r^{e_r}$。

**Power-of-two 子類**：$m = 2^k$ 時 $\Phi_m(x) = x^{2^{k-1}}+1$，$n = \varphi(m) = 2^{k-1}$；$m=8 \Rightarrow \Phi_8 = x^4+1$、$m=1024 \Rightarrow \Phi_{1024} = x^{512}+1$（Ring-LWE 實作首選）。

**Trace, Norm, Discriminant**：$\text{Tr}_{K/\mathbb{Q}}(\alpha) = \sum_i \sigma_i(\alpha)$、$N_{K/\mathbb{Q}}(\alpha) = \prod_i \sigma_i(\alpha)$；codifferent $\mathfrak{d}_{K/\mathbb{Q}}^{-1} = \{\alpha \in K : \text{Tr}(\alpha\mathcal{O}_K) \subseteq \mathbb{Z}\}$；判別式 $\Delta_K = N(\mathfrak{d}_{K/\mathbb{Q}})$。對 $m=2^k$：$|\Delta_K| = 2^{n(k-1)}$。

**Canonical (Minkowski) Embedding**：分圓域 $r_1=0,\ r_2=n/2$：

$$\sigma: K \hookrightarrow K_\mathbb{R} := K \otimes_\mathbb{Q} \mathbb{R} \cong \mathbb{C}^{n/2} \cong \mathbb{R}^n, \quad \alpha \mapsto (\sqrt{2}\,\text{Re}\,\sigma_1(\alpha),\ \sqrt{2}\,\text{Im}\,\sigma_1(\alpha),\ \ldots)$$

每個非零理想 $\mathfrak{a} \subseteq \mathcal{O}_K$ 在 $\sigma$ 下成為滿秩格（理想格），其覆積為：

$$\det(\sigma(\mathfrak{a})) = N(\mathfrak{a}) \cdot \sqrt{|\Delta_K|}$$

關鍵翻譯：理想（代數）$\leftrightarrow$ 滿秩格（幾何）；範數 $\leftrightarrow$ 覆積。Ring-LWE 攻擊難度即建基於理想格上的 SVP 仍困難。此節對應課程 Obj. 1–4：number rings、Dedekind ideals、prime splitting、norms、discriminants。

**質理想分裂（Obj. 3）**：對 $p \nmid m$，令 $f = \text{ord}_m(p) \in (\mathbb{Z}/m)^*$，則 $p\mathcal{O}_K = \mathfrak{p}_1 \cdots \mathfrak{p}_g$，$g = n/f$。

**Ring-LWE 關鍵選擇**：$q$ 質數且 $q \equiv 1 \pmod{m}$ $\Rightarrow$ $f=1$，$q$ 在 $\mathcal{O}_K$ 完全分裂：

$$R_q := \mathcal{O}_K / q\mathcal{O}_K \cong \prod_{i=1}^n \mathbb{Z}_q \quad \text{(CRT)}$$

此即 **NTT（Number-Theoretic Transform）** 的代數來源——一個課程目標 = 一個演算法加速。

**Worked Example（$m=8,\ n=4$）**：$K = \mathbb{Q}(\zeta_8)$，$|\Delta_K|=256$，$\sqrt{|\Delta_K|}=16$。選 $q=17$（因 $17 \equiv 1 \pmod{8}$）：$f=1$、$g=4$，$17\mathcal{O}_K = \mathfrak{q}_1\mathfrak{q}_2\mathfrak{q}_3\mathfrak{q}_4$ 完全分裂，$R_{17} \cong \mathbb{Z}_{17}^4$，即 NTT 結構。

---

## Step 3 — LWE and Its Ring Variant

**LWE（Regev, 2005）**：參數 $n,q \in \mathbb{Z}^+$、誤差分布 $\chi$（typically discrete Gaussian，寬度 $\alpha q$，$\alpha < 1/\sqrt{n}$）、秘密 $s \in \mathbb{Z}_q^n$；樣本為

$$(a,\ \langle a,s\rangle + e \bmod q) \in \mathbb{Z}_q^n \times \mathbb{Z}_q, \quad a \leftarrow U(\mathbb{Z}_q^n),\ e \leftarrow \chi.$$

直觀：求解帶雜訊的線性方程組；無雜訊時可用高斯消去 $O(n^3)$；加上小雜訊則指數困難（含量子）。Search-LWE（還原 $s$）與 Decision-LWE（區分均勻）多項式時間等價。

> **Regev 量子歸約**：任意 $n$ 維格上的 worst-case $\text{GapSVP}_\gamma$、$\text{SIVP}_\gamma$（$\gamma = \tilde{O}(n/\alpha)$）可規約至 average-case LWE——最壞情形困難性轉移至平均情形。**缺點**：公鑰 $A \in \mathbb{Z}_q^{n\times m}$ 大小 $O(n^2)$、乘法 $O(n^2)$。

**Ring-LWE（Lyubashevsky–Peikert–Regev, 2010）**：把 $\mathbb{Z}^n$ 換成分圓整數環

$$R := \mathcal{O}_K = \mathbb{Z}[x]/\Phi_m(x) \xrightarrow{m=2^k} \mathbb{Z}[x]/(x^n+1), \quad R_q := R/qR, \quad R^\vee := \mathfrak{d}_{K/\mathbb{Q}}^{-1}.$$

樣本為 $(a,\ as+e \bmod qR^\vee) \in R_q \times R^\vee_q$，$a \leftarrow U(R_q)$，$e \leftarrow \text{Gaussian on } K_\mathbb{R}$。對 power-of-two $m=2^k$：$R^\vee = (1/n)R$ 為純量倍，簡化為 simplified Ring-LWE。**歸約**：Ring-LWE $\Leftarrow$ approx-SVP on ideal lattices in $R$（量子）。

**NTT 與 Negacyclic 卷積**：$q \equiv 1 \pmod{m}$ $\Rightarrow$ $\Phi_m(x) \equiv \prod_{i=1}^n(x-\omega_i) \pmod{q}$，由 CRT 得 $R_q \cong \prod_{i=1}^n \mathbb{Z}_q$，計算 $O(n\log n)$。因 $x^n \equiv -1$，乘法為 negacyclic 卷積：

$$(a * b)_k = \sum_{i+j=k} a_i b_j - \sum_{i+j=k+n} a_i b_j, \quad 0 \leq k < n.$$

**LWE vs Ring-LWE：**

| 項目 | LWE | Ring-LWE |
|---|---|---|
| 公鑰大小 | $O(n^2)$ | $O(n)$（factor $n$） |
| 乘法複雜度 | $O(n^2)$ | $O(n\log n)$ via NTT |
| 安全歸約 | any-lattice SVP | ideal-lattice SVP |
| 結構 | 純線性代數 | 代數結構（環） |

Ring-LWE 用「結構」換「效率」——但結構也可能被攻擊利用，這是仍在研究的開放議題。

---

## Step 4 — Applications and Open Problems

**為什麼這件事重要**：今日加密生態完全依賴 RSA / ECC——HTTPS、行動銀行、即時通訊、軟體更新、VPN 皆然，量子電腦上線後將「同時」失效。Harvest-Now, Decrypt-Later 威脅意味醫療紀錄、外交電文、長期密鑰已暴露。NIST 建議 2030 前完成關鍵系統 PQC 轉換；Google、Cloudflare、Apple 已開始於 TLS 部署 ML-KEM 混合模式。

**NIST 後量子標準（2024）：**

- **FIPS 203 — ML-KEM (Kyber)**：金鑰交換 / KEM；公鑰約 800 byte、密文約 1 KB；快於 RSA。
- **FIPS 204 — ML-DSA (Dilithium)**：數位簽章；簽章約 2.4 KB、公鑰約 1.3 KB。
- Keys < 2 KB；安全性 $\approx$ AES-128。
- 數學基礎：Ring-LWE / Module-LWE。

**Open Problems：**

- Quantum hardness of Ideal-SVP vs SVP——理想格是否真比一般格容易？
- Subfield / Galois attacks via $\text{Gal}(K/\mathbb{Q})$：利用代數結構加速攻擊。
- Security of non-cyclotomic rings：NTRU Prime 等變體。
- Classical (non-quantum) worst-case 歸約仍開放。

**問題鏈（The Problem Chain）：**

$$\text{RSA/ECC} \to \text{整數分解/ECDLP} \xrightarrow{\text{Shor 1994}} \text{量子多項式破解} \to \text{Lattice/SVP}_\gamma\text{（量子難解）}$$
$$\to \text{LWE (2005，公鑰 }O(n^2)\text{)} \to \textbf{Ring-LWE (2010，公鑰 }O(n)\text{、NTT 加速 }O(n\log n)\text{)}$$

**Course Connections**：§2 number rings & discriminants (Obj. 1, 4) | §2 Dedekind ideals & norms (Obj. 2) | §3 prime splitting in $R_q$ (Obj. 3) | §2 Minkowski / geometry of numbers (Obj. 5)。

**結語**：本報告以代數數論的觀點理解 Ring-LWE，串連「格 → 理想格 → Ring-LWE → 後量子密碼學」。代數數論 + 格幾何 = 你我未來十年的網路安全基礎。

---

## 主要參考文獻

- O. Regev, *J. ACM* **56**(6), 2009.
- V. Lyubashevsky, C. Peikert, O. Regev, *J. ACM* **60**(6), 2013.
- C. Peikert, *Found. Trends TCS* **10**(4), 2016.
- NIST FIPS 203 / 204, 2024.
- D. A. Marcus, *Number Fields*, Ch. 1–5.

## 授權

MIT License — 見 [`LICENSE`](LICENSE)
