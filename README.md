# 數論（一）期末報告 — Ring-LWE 與理想格密碼學

> Algebraic Number Theory Meets Post-Quantum Security
>
> 黃崇晉 (Chung-Chin Huang) · L16141149 · 數論（一）· 2026 春季 · 國立成功大學

## 報告主題
以代數數論的觀點理解 Ring-LWE：串連「格 → 理想格 → Ring-LWE → 後量子密碼學」。

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

## 報告章節
1. **Motivation** — 為什麼需要後量子密碼學 (RSA / ECC、Shor、Harvest-Now-Decrypt-Later)
2. **Step 1** — 格與幾何數論 (Lattice, SVP, CVP, Minkowski 第一定理)
3. **Step 2** — 分圓域與理想格 (Cyclotomic Fields, Ideal Lattices, Prime Splitting)
4. **Step 3** — LWE 與 Ring-LWE (Regev 量子歸約, NTT, Negacyclic Convolution)
5. **Step 4** — 應用與開放問題 (NIST FIPS 203/204, ML-KEM, ML-DSA)

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

## 主要參考文獻
- O. Regev, *J. ACM* **56**(6), 2009.
- V. Lyubashevsky, C. Peikert, O. Regev, *J. ACM* **60**(6), 2013.
- C. Peikert, *Found. Trends TCS* **10**(4), 2016.
- NIST FIPS 203 / 204, 2024.
- D. A. Marcus, *Number Fields*, Ch. 1–5.

## 授權
MIT License — 見 [`LICENSE`](LICENSE)
