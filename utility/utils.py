# utils.py 或 utility/tournament_utils.py

def to_chinese_match_no(n):
    """
    將場次數字轉換為指定國字格式：
    - 1-9: 一, 二...
    - 10-19: 十, 十一...
    - 20, 30...: 二十, 三十...
    - 21, 35...: 二一, 三五 (不含十)
    - 超過 99: 回傳 '不支援'
    """
    if n <= 0:
        return ""
    if n > 99:
        return "不支援"
        
    # 國字定義
    units = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"]
    
    # 1. 個位數 (1-9)
    if n < 10:
        return units[n]
    
    # 2. 十位數為 1 的處理 (10-19)
    elif n < 20:
        if n == 10:
            return "十"
        else:
            return "十" + units[n % 10]
            
    # 3. 20-99 的處理
    else:
        tens = n // 10
        remainder = n % 10
        
        # 剛好是整十 (20, 30, 40...)
        if remainder == 0:
            return units[tens] + "十"
        # 21, 35, 99 格式 (二一, 三五)
        else:
            return units[tens] + units[remainder]

# 這裡未來可以加入其他共用工具，例如：
# def format_team_name(name): ...
# def validate_excel_format(file): ...
