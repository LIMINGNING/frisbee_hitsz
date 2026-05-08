import csv
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent
CSV_PATH = BASE_DIR / "dingdan2026.csv"
OUTPUT_PATH = BASE_DIR / "latex_teams.txt"
MAIN_EVENT_MARKER = "5月16日和23日正式比赛"

TEAM_ASSIGNMENTS = {
    "A队": ["康钊源", "赵浩帆", "寇家辉", "王俊杰", "吴智杰", "罗梦恬", "潘秀丽", "樊黎硕", "肖叶萱"],
    "B队": ["刘慕实", "马荣", "梁柱政", "杨鸿鑫", "金圣博", "Ocean", "陈子瑶", "杨千凝", "孙明琪"],
    "C队": ["黄家乐", "黄家骏", "金珉宇", "典珅司", "刘沐之", "小双", "EL MOUDDEN ZAYNAB", "李焱", "小午"],
    "D队": ["王硕", "耿增越", "温岩", "万一鸣", "张梦辉", "孟欣", "段绚涵", "summer", "刘昊媛"],
}


def load_signup_rows():
    with CSV_PATH.open(encoding="utf-8-sig", newline="") as f:
        records = list(csv.reader(f))

    header = records[2]
    return [dict(zip(header, row)) for row in records[3:]]


def normalize_gender(value):
    text = str(value or "").strip().lower()
    if "female" in text or "女" in text:
        return "女"
    if "male" in text or "男" in text:
        return "男"
    return "未知"


def experience_score(row):
    name = row["姓名 Name"].strip()
    text = row.get("盘龄 Frisbee experience") or ""

    if name == "黄家乐":
        return 5
    if "2年<t" in text:
        return 4
    if "1年<t<2年" in text:
        return 3
    if "6月<t<1年" in text:
        return 2
    if "t<6月" in text:
        return 1
    return 0


def experience_label(row):
    text = row.get("盘龄 Frisbee experience") or ""

    if "2年<t" in text:
        return "2年以上"
    if "1年<t<2年" in text:
        return "1-2年"
    if "6月<t<1年" in text:
        return "6月-1年"
    if "t<6月" in text:
        return "6月以内"
    return "-"


def latex_escape(value):
    text = str(value or "").strip()
    return text.replace("&", r"\&")


def validate_teams(rows_by_name):
    assigned = [name for names in TEAM_ASSIGNMENTS.values() for name in names]
    duplicates = sorted({name for name in assigned if assigned.count(name) > 1})
    missing = sorted(set(rows_by_name) - set(assigned))
    extra = sorted(set(assigned) - set(rows_by_name))

    if duplicates or missing or extra:
        raise ValueError(f"分队名单不匹配: duplicates={duplicates}, missing={missing}, extra={extra}")


def build_latex(rows_by_name):
    parts = [r"\appendix", r"\section{参赛人员名单}", ""]

    for team_name, names in TEAM_ASSIGNMENTS.items():
        parts.extend(
            [
                rf"\subsection*{{{team_name}}}",
                r"\renewcommand{\arraystretch}{1.2}",
                r"\begin{tabularx}{\textwidth}{@{} L{3cm} M{1.4cm} L{3cm} X @{}}",
                r"    \toprule",
                r"    \textbf{姓名} & \textbf{性别} & \textbf{学号} & \textbf{盘龄} \\",
                r"    \midrule",
            ]
        )

        for name in names:
            row = rows_by_name[name]
            display_name = latex_escape(row["姓名 Name"])
            gender = normalize_gender(row["性别 Gender"])
            student_id = latex_escape(row.get("学号 student ID") or "-")
            frisbee_experience = experience_label(row)
            parts.append(f"    {display_name} & {gender} & {student_id} & {frisbee_experience} \\\\")

        parts.extend([r"    \bottomrule", r"\end{tabularx}", ""])

    return "\n".join(parts)


def print_balance(rows_by_name):
    for team_name, names in TEAM_ASSIGNMENTS.items():
        rows = [rows_by_name[name] for name in names]
        men = sum(normalize_gender(row["性别 Gender"]) == "男" for row in rows)
        women = sum(normalize_gender(row["性别 Gender"]) == "女" for row in rows)
        total_score = sum(experience_score(row) for row in rows)
        experienced = sum(experience_score(row) > 1 for row in rows)
        print(
            f"{team_name}: {len(rows)}人, 男/女={men}/{women}, "
            f"盘龄均分={total_score / len(rows):.2f}, 老手/新手={experienced}/{len(rows) - experienced}"
        )


def main():
    rows = [row for row in load_signup_rows() if MAIN_EVENT_MARKER in (row.get("报名活动") or "")]
    rows_by_name = {row["姓名 Name"].strip(): row for row in rows}

    validate_teams(rows_by_name)
    OUTPUT_PATH.write_text(build_latex(rows_by_name), encoding="utf-8")
    print_balance(rows_by_name)
    print(f"Wrote {OUTPUT_PATH}")


if __name__ == "__main__":
    main()
