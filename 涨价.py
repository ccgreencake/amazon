def get_A(x):
    """
    根据 x（原售价）返回原抽佣利率 A（小数）。
    当 x < 15 时 A=5%；15 ≤ x < 20 时 A=10%；x ≥ 20 时 A=17%。
    """
    if x < 15:
        return 0.05
    elif x < 20:
        return 0.10
    else:
        return 0.17

def get_B(x_plus_y):
    """
    根据 x+y（调价后售价总和）返回当前抽佣利率 B（小数）。
    规则同 A 的分段。
    """
    if x_plus_y < 15:
        return 0.05
    elif x_plus_y < 20:
        return 0.10
    else:
        return 0.17

def compute_y_scheme1(x, C, D, tol=1e-8, max_iter=1000):
    """
    方案 1: y = x*(B - A + 0.15*C) / (1 - B - 0.15*C - D)
    隐式地通过固定点迭代求解。
    """
    y = 0.0
    for i in range(max_iter):
        A = get_A(x)
        B = get_B(x + y)
        num = x * (B - A + 0.15 * C)
        den = 1 - B - 0.15 * C - D
        if abs(den) < 1e-12:
            raise ZeroDivisionError("方案1分母接近零，可能无解或不稳定。")
        y_new = num / den
        if abs(y_new - y) < tol:
            return y_new
        y = y_new
    raise RuntimeError(f"方案1未在 {max_iter} 次迭代内收敛，请检查参数或增大 max_iter。")

def compute_y_scheme2(x, C, tol=1e-8, max_iter=1000):
    """
    方案 2: y = x*(B - A + 0.15*C) / (1 - B - 0.15*C)
    同样通过固定点迭代求解（因 B 仍依赖 x+y）。
    """
    y = 0.0
    for i in range(max_iter):
        A = get_A(x)
        B = get_B(x + y)
        num = x * (B - A + 0.15 * C)
        den = 1 - B - 0.15 * C
        if abs(den) < 1e-12:
            raise ZeroDivisionError("方案2分母接近零，可能无解或不稳定。")
        y_new = num / den
        if abs(y_new - y) < tol:
            return y_new
        y = y_new
    raise RuntimeError(f"方案2未在 {max_iter} 次迭代内收敛，请检查参数或增大 max_iter。")

if __name__ == "__main__":
    # 1）读取基础参数
    x = float(input("请输入原售价 x（数字）："))
    C = float(input("请输入关税率 C（如 0.1 表示 10%）："))
    D = float(input("请输入原利润率 D（如 0.2 表示 20%）："))

    # 2）选择方案
    print("\n请选择涨价方案：")
    print("  1. 方案1：y = x*(B - A + 0.15*C) / (1 - B - 0.15*C - D)")
    print("  2. 方案2：y = x*(B - A + 0.15*C) / (1 - B - 0.15*C)")
    choice = input("输入 1 或 2：").strip()

    try:
        if choice == "1":
            y = compute_y_scheme1(x, C, D)
            scheme = "方案1"
        elif choice == "2":
            y = compute_y_scheme2(x, C)
            scheme = "方案2"
        else:
            raise ValueError("无效的方案选择，仅支持输入“1”或“2”。")

        print(f"\n{scheme} 计算结果：")
        print(f"  需要涨价 y ≈ {y:.4f}")
        print(f"  涨价后售价 x+y ≈ {x + y:.4f}")
        print(f"  原抽佣 A = {get_A(x)*100:.1f}%")
        print(f"  新抽佣 B = {get_B(x+y)*100:.1f}%")
    except Exception as e:
        print(f"计算失败：{e}")
