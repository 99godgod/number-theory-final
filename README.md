# 數論（一）期末報告 — Ring-LWE 與理想格密碼學

> 幾何數論、分圓域、Dedekind 整環與後量子安全基礎
>
> 黃崇晉 (Chung-Chin Huang) · L16141149 · 數論（一）· 2026 春季 · 國立成功大學

## 檔案結構

| 檔案 | 說明 |
|---|---|
| `summary/summary.pdf` / `summary/summary.tex` | 書面摘要（含 Minkowski 定理完整證明、分圓域與理想格、Ring-LWE 代數結構） |
| `presentation/presentation.pdf` / `.tex` | 中文投影片（beamer，20 頁，約 40 分鐘）：前置六頁簡介格／Minkowski／分圓域／理想格／SVP／CVP，重點在 LWE 與 Ring-LWE，末段 5 分鐘談應用與開放問題 |
| `presentation/presentation_en.pdf` / `.tex` | 英文投影片（同上之 English 版，20 頁） |
| `presentation/speaker_script.pdf` / `.tex` | 口語報告提示稿（中英對照，7 頁），逐頁對應 `presentation/presentation.pdf` 全 20 頁 |
| `summary/ring_lwe_summary.pdf` / `summary/ring_lwe_summary_zh.pdf` | 早期一頁式精簡大綱草稿（英文 / 中文對照版） |
| `report_excerpt.pdf` / `report_excerpt.tex` | 外部資料節錄（繁中，2 頁）：自 HackMD 筆記與 Drive 講義中，僅節錄與報告主題貼合之段落並對應各章節 |

## 編譯方式

LaTeX 來源以 `xelatex` 編譯（需 xeCJK 與 Microsoft JhengHei 字型）；投影片與提示稿含頁碼／章節交叉參照，須各跑兩趟：
```bash
(cd summary && xelatex summary.tex)
(cd presentation && xelatex presentation.tex && xelatex presentation.tex)
(cd presentation && xelatex presentation_en.tex && xelatex presentation_en.tex)
(cd presentation && xelatex speaker_script.tex && xelatex speaker_script.tex)
```

---

## Section 1 — 幾何數論基礎：格與 Minkowski 定理

**格的定義**：取 $`\mathbb{R}^n`$ 中 $`n`$ 個線性獨立向量 $`b_1,\ldots,b_n`$，整數線性組合

$$\mathcal{L} = \sum_{i=1}^{n} \mathbb{Z}\, b_i \subset \mathbb{R}^n$$

稱為秩 $`n`$ 的格。基底矩陣 $`B=[b_1|\cdots|b_n]`$ 的行列式絕對值 $`\det(\mathcal{L}):=|\det B|`$ 稱為**覆積**，與基底選擇無關。

**Blichfeldt 引理（1914）**：設 $`\mathcal{L}\subset\mathbb{R}^n`$ 為格，$`S\subset\mathbb{R}^n`$ 為 Lebesgue 可測集。若 $`\mathrm{vol}(S)>\det(\mathcal{L})`$，則存在相異 $`x,y\in S`$ 使得 $`x-y\in\mathcal{L}`$。

> 證明：令 $`\mathcal{F}`$ 為 $`\mathcal{L}`$ 的基本域，對 $`\lbrace S_v := S\cap(v+\mathcal{F})\rbrace_{v\in\mathcal{L}}`$ 做平移後，$`\sum\mathrm{vol}(S_v') = \mathrm{vol}(S) > \mathrm{vol}(\mathcal{F})`$，由鴿巢原理得兩個平移片重疊，差向量即為所求。

**Minkowski 第一定理（1896）**：設 $`S\subset\mathbb{R}^n`$ 為中心對稱凸體。若 $`\mathrm{vol}(S)>2^n\det(\mathcal{L})`$，則 $`S`$ 包含至少一個非零格點。

> 證明：對 $`\tfrac{1}{2}S`$ 套用 Blichfeldt 引理，利用中心對稱性與凸性知差向量 $`p-q\in S\cap\mathcal{L}\setminus\lbrace 0\rbrace`$。

**Minkowski 界（推論）**：設 $`K`$ 為次數 $`n=r_1+2r_2`$ 的數域，$`\Delta_K`$ 為判別式，則 $`\mathrm{Cl}(K)`$ 中每個理想類均含範數有界的整數理想：

$$N(\mathfrak{a})\;\leq\;M_K\;:=\;\sqrt{|\Delta_K|}\cdot\Bigl(\tfrac{4}{\pi}\Bigr)^{r_2}\cdot\frac{n!}{n^n}$$

此保證類群 $`\mathrm{Cl}(K)`$ 有限（類數有限定理）。Minkowski 定理的**非構造性**連結幾何（格體積）與算術（理想範數）：它保證格點*存在*，卻不給出尋找方法——這既是 SVP 困難性的根源，也是密碼學安全性的幾何基石。

---

## Section 2 — 分圓域與理想格結構

**分圓域設定**：令 $`n=2^k`$，$`\Phi_{2n}(x)=x^n+1`$（$`n=2^k`$ 時在 $`\mathbb{Q}`$ 上不可約），$`K=\mathbb{Q}(\zeta_{2n})`$，$`[K:\mathbb{Q}]=n`$，$`r_1=0`$，$`r_2=n/2`$。

**Dedekind 整環**：$`\mathcal{O}_K=\mathbb{Z}[\zeta]`$ 是 Dedekind 整環（Noetherian、整閉、每個非零質理想皆極大），因此每個非零理想具有唯一的質理想積分解：$`\mathfrak{a}=\mathfrak{p}_1^{e_1}\cdots\mathfrak{p}_r^{e_r}`$。

**標準嵌入（Canonical Embedding）**：$`K`$ 有 $`n`$ 個複數嵌入 $`\sigma_j(\zeta)=\zeta^{2j-1}`$，$`j=1,\ldots,n`$。共軛對稱使 $`\sigma`$ 的像落在同構於 $`\mathbb{R}^n`$ 的子空間 $`H\subset\mathbb{C}^n`$ 中，且環乘法與逐分量乘法相容。

**判別式與格行列式**：$`\sigma(\mathcal{O}_K)\subset H`$ 為秩 $`n`$ 的格，對 $`K=\mathbb{Q}(\zeta_{2n})`$（$`n=2^k`$）：

$$\det\bigl(\sigma(\mathcal{O}_K)\bigr)=\sqrt{|\Delta_K|}=n^{n/2}$$

**理想格與範數**：$`\mathcal{O}_K`$ 的分式理想 $`\mathfrak{a}`$ 透過標準嵌入給出理想格 $`\sigma(\mathfrak{a})`$，覆積為

$$\det\bigl(\sigma(\mathfrak{a})\bigr)=N(\mathfrak{a})\cdot\sqrt{|\Delta_K|}$$

**差異理想與逆差理想**：$`\mathfrak{D}_{K/\mathbb{Q}}=(n\zeta^{n-1})=(n)`$（因 $`\zeta^{n-1}`$ 為單位元），逆差理想為

$$\mathfrak{D}^{-1}=\bigl\lbrace x\in K:\mathrm{Tr}_{K/\mathbb{Q}}(x\,\mathcal{O}_K)\subset\mathbb{Z}\bigr\rbrace=\tfrac{1}{n}\mathcal{O}_K$$

$`\mathfrak{D}^{-1}`$ 是 $`\mathcal{O}_K`$ 在跡配對 $`\langle a,b\rangle=\mathrm{Tr}(ab)`$ 下的**對偶格**。

---

## Section 3 — Ring-LWE 的數論結構

**問題設定**：令 $`R=\mathcal{O}_K=\mathbb{Z}[x]/\Phi_{2n}(x)`$，$`R_q=R/qR\cong\mathbb{Z}_q[x]/\Phi_{2n}(x)`$。

**Ring-LWE 假設（LPR 2010）**：對秘密 $`s\in R_q`$ 與離散 Gaussian 誤差分布 $`\chi_R`$，分布 $`(a,\;a\cdot s+e)`$（$`a\overset{\$}{\leftarrow}R_q`$，$`e\leftarrow\chi_R`$）與 $`R_q\times R_q`$ 均勻分布計算不可區分。量子歸約：worst-case ideal-SVP$`_\gamma`$ → average-case Ring-LWE。

**質理想分解**：對質數 $`q\nmid 2n`$，$`\Phi_{2n}(x)\bmod q`$ 的因式分解 $`\prod_{i=1}^g\phi_i(x)^{e_i}`$（$`\deg\phi_i=f`$，$`efg=n`$）決定 $`R_q`$ 的結構。**完全分裂**（$`e=1,f=1,g=n`$）當且僅當 $`q\equiv 1\pmod{2n}`$，此時由中國剩餘定理：

$$R_q\;\cong\;\mathbb{F}_q^n$$

**NTT 的數論根源**：完全分裂條件使 $`R_q`$ 的乘法化為逐分量乘法，評估映射 $`a\mapsto(\hat{a}(\omega_1),\ldots,\hat{a}(\omega_n))`$ 即為**數論變換（NTT）**，以 Cooley-Tukey 蝴蝶運算在 $`O(n\log n)`$ 時間完成（優於直接乘法的 $`O(n^2)`$）。

**誤差的自然定義域**：嚴格的 Ring-LWE 表述中，誤差定義在 $`\mathfrak{D}^{-1}/q\mathfrak{D}^{-1}=\frac{1}{n}R_q`$ 而非 $`R_q`$，確保誤差分布在跡配對下具代數不變性與安全歸約的嚴密性。

---

## Section 4 — 應用與延伸

**NIST 後量子標準（2024 年 8 月）：**

| 標準 | 算法 | 基礎 | 參數（Level I） |
|---|---|---|---|
| FIPS 203 | ML-KEM (Kyber) | Module-LWE | $`n=256`$，$`q=3329`$，攻擊難度 $`\approx 2^{118}`$ |
| FIPS 204 | ML-DSA (Dilithium) | Module-LWE + Module-SIS | 公鑰約 1.3 KB，簽章約 2.4 KB |

安全鏈：破解密碼 $`\Leftrightarrow`$ 解 BDD on ideal lattice $`\Leftarrow`$ ideal-SVP$`_\gamma`$ 困難。

**開放問題：**

- **Ideal-SVP vs SVP**：理想格的 SVP 是否真的不比任意格容易？Cramer et al.（2016）對特定分圓環給出量子加速，但一般情形仍開放。
- **非分圓域的 Ring-LWE**：LPR（2010）的歸約依賴分圓域的交換 Galois 結構；對 Galois 閉包為非交換群的數環，安全歸約是否成立仍未知。
- **古典 worst-case 歸約**：Regev 的 LWE 歸約為量子的；純古典 worst-case → average-case 歸約至今未找到。
- **SVP 量子下界**：最佳量子演算法為 $`2^{0.265n}`$（Laarhoven 2015），但無條件指數下界至今未被證明。

---

## 主要參考文獻

- O. Regev, *J. ACM* **56**(6), 2009.
- V. Lyubashevsky, C. Peikert, O. Regev, *J. ACM* **60**(6), 2013.
- H. Minkowski, *Geometrie der Zahlen*, 1896.
- H. Blichfeldt, *Trans. AMS* **15**(3), 1914.
- J. Neukirch, *Algebraic Number Theory*, Springer, 1999.
- NIST FIPS 203 / 204, 2024.

## 授權

MIT License — 見 [`LICENSE`](LICENSE)
