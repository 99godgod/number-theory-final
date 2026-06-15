"""
NTT 多項式乘法範例
==================
示範以數論轉換（Number Theoretic Transform, NTT）加速多項式乘法，
複雜度由樸素的 O(n^2) 降至 O(n log n)，全程整數運算、零浮點誤差。

包含兩種卷積：
  1. 循環卷積 (cyclic)      —— 在環 Z_q[x]/(x^n - 1) 下，使用 n 次單位根 omega
  2. 負循環卷積 (negacyclic) —— 在環 Z_q[x]/(x^n + 1) 下（Ring-LWE 所用），
                               使用 2n 次單位根 psi（psi^2 = omega）

模數需滿足 q 為質數且 q ≡ 1 (mod 2n)，以保證單位根存在。
本例使用 q = 7681（= 2^9 * 15 + 1），支援 n 為 2 的冪且 n <= 256。
"""

# ──────────────────────────────────────────────
# 基礎數論工具
# ──────────────────────────────────────────────
def is_prime(num):
    if num < 2:
        return False
    i = 2
    while i * i <= num:
        if num % i == 0:
            return False
        i += 1
    return True

def primitive_root(q):
    """求質數 q 的一個原根 (primitive root)。"""
    # 對 q-1 做質因數分解
    phi = q - 1
    factors, m, d = [], phi, 2
    while d * d <= m:
        if m % d == 0:
            factors.append(d)
            while m % d == 0:
                m //= d
        d += 1
    if m > 1:
        factors.append(m)
    # 測試候選 g 是否為原根：對每個質因數 p，g^(phi/p) != 1
    for g in range(2, q):
        if all(pow(g, phi // p, q) != 1 for p in factors):
            return g
    raise ValueError("找不到原根")

def bit_reverse(a):
    """就地位元反轉重排 (in-place bit-reversal permutation)。"""
    n = len(a)
    j = 0
    for i in range(1, n):
        bit = n >> 1
        while j & bit:
            j ^= bit
            bit >>= 1
        j ^= bit
        if i < j:
            a[i], a[j] = a[j], a[i]

# ──────────────────────────────────────────────
# NTT 核心（迭代式 Cooley-Tukey）
# ──────────────────────────────────────────────
def ntt(a, q, root):
    """正向 NTT。root 為 n 次本原單位根。就地運算。"""
    n = len(a)
    bit_reverse(a)
    length = 2
    while length <= n:
        # w_len = root^(n/length) 為 length 次本原單位根
        wlen = pow(root, n // length, q)
        for i in range(0, n, length):
            w = 1
            for j in range(length // 2):
                u = a[i + j]
                v = a[i + j + length // 2] * w % q
                a[i + j] = (u + v) % q
                a[i + j + length // 2] = (u - v) % q
                w = w * wlen % q
        length <<= 1
    return a

def intt(a, q, root):
    """逆向 NTT，使用 root 的反元素並乘上 n 的反元素。"""
    n = len(a)
    root_inv = pow(root, q - 2, q)          # 費馬小定理求反元素
    ntt(a, q, root_inv)
    n_inv = pow(n, q - 2, q)
    return [x * n_inv % q for x in a]

# ──────────────────────────────────────────────
# 多項式乘法（NTT 版）
# ──────────────────────────────────────────────
def poly_mul_cyclic(a, b, q):
    """循環卷積：結果在 Z_q[x]/(x^n - 1) 下，n 為 a、b 補零後的長度。"""
    n = len(a)
    assert len(b) == n and (n & (n - 1)) == 0, "長度需相等且為 2 的冪"
    assert (q - 1) % n == 0, "需 q ≡ 1 (mod n)"
    g = primitive_root(q)
    omega = pow(g, (q - 1) // n, q)         # n 次本原單位根
    fa, fb = ntt(a[:], q, omega), ntt(b[:], q, omega)
    fc = [x * y % q for x, y in zip(fa, fb)]
    return intt(fc, q, omega)

def poly_mul_negacyclic(a, b, q):
    """負循環卷積：結果在 Ring-LWE 環 Z_q[x]/(x^n + 1) 下。"""
    n = len(a)
    assert len(b) == n and (n & (n - 1)) == 0, "長度需相等且為 2 的冪"
    assert (q - 1) % (2 * n) == 0, "需 q ≡ 1 (mod 2n)"
    g = primitive_root(q)
    psi = pow(g, (q - 1) // (2 * n), q)     # 2n 次本原單位根
    psi_inv = pow(psi, q - 2, q)
    omega = psi * psi % q                    # n 次本原單位根
    # 前置權重：a_i <- a_i * psi^i
    aw = [a[i] * pow(psi, i, q) % q for i in range(n)]
    bw = [b[i] * pow(psi, i, q) % q for i in range(n)]
    fa, fb = ntt(aw, q, omega), ntt(bw, q, omega)
    fc = [x * y % q for x, y in zip(fa, fb)]
    c = intt(fc, q, omega)
    # 後置權重：c_i <- c_i * psi^-i
    return [c[i] * pow(psi_inv, i, q) % q for i in range(n)]

# ──────────────────────────────────────────────
# 樸素參考實作（O(n^2)），用於驗證正確性
# ──────────────────────────────────────────────
def naive_cyclic(a, b, q):
    n = len(a)
    c = [0] * n
    for i in range(n):
        for j in range(n):
            c[(i + j) % n] = (c[(i + j) % n] + a[i] * b[j]) % q
    return c

def naive_negacyclic(a, b, q):
    n = len(a)
    c = [0] * n
    for i in range(n):
        for j in range(n):
            k = i + j
            if k < n:
                c[k] = (c[k] + a[i] * b[j]) % q
            else:                            # x^n ≡ -1，超出部分變號折回
                c[k - n] = (c[k - n] - a[i] * b[j]) % q
    return c

# ──────────────────────────────────────────────
# 示範
# ──────────────────────────────────────────────
def poly_str(p):
    terms = []
    for i, c in enumerate(p):
        if c == 0:
            continue
        terms.append(f"{c}" if i == 0 else f"{c}x^{i}")
    return " + ".join(terms) if terms else "0"

if __name__ == "__main__":
    q = 7681
    n = 8
    assert is_prime(q) and (q - 1) % (2 * n) == 0

    # 範例多項式（係數低次到高次）
    a = [1, 2, 3, 4, 0, 0, 0, 0]
    b = [5, 6, 7, 8, 0, 0, 0, 0]

    print(f"模數 q = {q}（質數: {is_prime(q)}），維度 n = {n}")
    print(f"原根 g = {primitive_root(q)}")
    print(f"a(x) = {poly_str(a)}")
    print(f"b(x) = {poly_str(b)}")
    print("-" * 60)

    # 循環卷積 (x^n - 1)
    c_ntt = poly_mul_cyclic(a, b, q)
    c_ref = naive_cyclic(a, b, q)
    print("[循環卷積  Z_q[x]/(x^n - 1)]")
    print(f"  NTT  : {c_ntt}")
    print(f"  樸素 : {c_ref}")
    print(f"  一致 : {c_ntt == c_ref}")
    print("-" * 60)

    # 負循環卷積 (x^n + 1) —— Ring-LWE
    d_ntt = poly_mul_negacyclic(a, b, q)
    d_ref = naive_negacyclic(a, b, q)
    print("[負循環卷積  Z_q[x]/(x^n + 1)  ← Ring-LWE]")
    print(f"  NTT  : {d_ntt}")
    print(f"  樸素 : {d_ref}")
    print(f"  一致 : {d_ntt == d_ref}")
    print("-" * 60)

    # 較大規模隨機測試
    import random
    n2 = 256
    q2 = 7681                                # 7680 = 2^9 * 15，2*256=512 | 7680
    assert (q2 - 1) % (2 * n2) == 0
    ra = [random.randrange(q2) for _ in range(n2)]
    rb = [random.randrange(q2) for _ in range(n2)]
    ok = poly_mul_negacyclic(ra, rb, q2) == naive_negacyclic(ra, rb, q2)
    print(f"隨機測試 n={n2}：負循環卷積 NTT == 樸素 → {ok}")
