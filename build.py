import os

# 罫線セットの定義
LINE_SETS = {
    "thin": {
        "0,0,1,1": "─", "0,1,0,1": "┌", "0,1,1,0": "┐", "0,1,1,1": "┬",
        "1,0,0,1": "└", "1,0,1,0": "┘", "1,0,1,1": "┴", "1,1,0,0": "│",
        "1,1,0,1": "├", "1,1,1,0": "┤", "1,1,1,1": "┼"
    },
    "thick": {
        "0,0,1,1": "━", "0,1,0,1": "┏", "0,1,1,0": "┓", "0,1,1,1": "┳",
        "1,0,0,1": "┗", "1,0,1,0": "┛", "1,0,1,1": "┻", "1,1,0,0": "┃",
        "1,1,0,1": "┣", "1,1,1,0": "┫", "1,1,1,1": "╋"
    }
}

# デフォルト値
DEFAULT_EMPTY = ""

def generate_linemode_code(style):
    lines = []
    lines.append("Dim linemode(1, 1, 1, 1)")
    for i in range(2):
        for j in range(2):
            for k in range(2):
                for l in range(2):
                    key = f"{i},{j},{k},{l}"
                    val = LINE_SETS[style].get(key, DEFAULT_EMPTY)
                    lines.append(f'linemode({i}, {j}, {k}, {l}) = "{val}"')
    return "\n".join(lines)

def get_defchar(direction, style):
    if style == "thin":
        return "│" if direction in ["Top", "Bottom"] else "─"
    else:
        return "┃" if direction in ["Top", "Bottom"] else "━"

def build():
    # テンプレートを読み込む (後ほど作成)
    with open("template.vbs", "r", encoding="shift_jis") as f:
        template = f.read()

    configs = [
        ("bottom_line.vbs", "Bottom", "thin"),
        ("bottom_line_b.vbs", "Bottom", "thick"),
        ("left_line.vbs", "Left", "thin"),
        ("left_line_b.vbs", "Left", "thick"),
        ("right_line.vbs", "Right", "thin"),
        ("right_line_b.vbs", "Right", "thick"),
        ("top_line.vbs", "Top", "thin"),
        ("top_line_b.vbs", "Top", "thick"),
    ]

    for filename, direction, style in configs:
        linemode_code = generate_linemode_code(style)
        defchar = get_defchar(direction, style)
        
        # テンプレートのプレースホルダーを置換
        content = template.replace("{{LINEMODE_CODE}}", linemode_code)
        content = content.replace("{{DIRECTION}}", f'"{direction}"')
        content = content.replace("{{DEFCHAR}}", f'"{defchar}"')
        
        with open(filename, "w", encoding="shift_jis") as f:
            f.write(content)
        print(f"Generated {filename}")

if __name__ == "__main__":
    build()
