def split_timetable(timetable: str, separator='/'):
    timetable_list = timetable.split(sep=separator)
    timetable_dic = {}
    for time in timetable_list:
        hour = time.split(":")[0]
        if hour not in timetable_dic.keys():
            timetable_dic[hour] = []
        timetable_dic[hour].append(time)
    return timetable_dic