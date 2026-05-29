import re

# 期待される罫線文字リスト
EXPECTED_JOINTS = {
    "─", "┌", "┐", "┬", "└", "┘", "┴", "│", "├", "┤", "┼",
    "━", "┏", "┓", "┳", "┗", "┛", "┻", "┃", "┣", "┫", "╋"
}

def test_template():
    with open("template.vbs", "r", encoding="cp932", errors="ignore") as f:
        content = f.read()
    
    # top_joint など 4 つの配列の定義を抽出
    match = re.search(r'top_joint\s+=\s+Array\((.*?)\)', content, re.DOTALL)
    if not match:
        print("ERROR: top_joint not found")
        return False
    
    # 文字列を抽出して集合化
    items_str = match.group(1)
    # クォートで囲まれた文字を取得
    found_joints = set(re.findall(r'"([^"]+)"', items_str))
    
    missing = EXPECTED_JOINTS - found_joints
    extra = found_joints - EXPECTED_JOINTS
    
    if missing:
        print(f"FAILED: Missing joints: {missing}")
        return False
    if extra:
        print(f"FAILED: Extra joints: {extra}")
        return False
    
    print("PASSED: All joints are correct.")
    return True

if __name__ == "__main__":
    if not test_template():
        exit(1)
