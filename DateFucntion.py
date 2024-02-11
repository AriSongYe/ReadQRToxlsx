from datetime import datetime, timedelta

def get_sunday(date_obj):
    # 입력된 날짜의 요일이 일요일(0)이면 해당 날짜를 반환
    if date_obj.weekday() == 6:
        return date_obj.replace(hour=0, minute=0, second=0, microsecond=0)  # 시간을 00:00:00으로 설정하여 datetime 객체 반환
    else:
        # 입력된 날짜가 일요일이 아닌 경우 해당 주의 일요일로 이동하여 반환
        days_until_sunday = date_obj.weekday() + 1  # 일요일까지의 날짜 수 계산
        sunday = date_obj - timedelta(days=days_until_sunday)
        return sunday.replace(hour=0, minute=0, second=0, microsecond=0)  # 시간을 00:00:00으로 설정하여 datetime 객체 반환

def get_current_sunday():
    # 현재 날짜를 가져옴
    now = datetime.now()
    # 현재 날짜의 일요일을 구함
    sunday = get_sunday(now)
    return sunday
