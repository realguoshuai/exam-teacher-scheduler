"""
工具函数
"""

import pandas as pd
import os
from typing import List, Dict, Any
from models import Teacher, Exam, Schedule
import config


def init_data_dir():
    """初始化数据目录"""
    if not os.path.exists(config.DATA_DIR):
        os.makedirs(config.DATA_DIR)
        print(f"创建数据目录: {config.DATA_DIR}")

        _create_sample_teachers()
        _create_sample_exams()
        _create_default_config()
    else:
        # 如果目录已存在，确保配置文件也存在
        if not os.path.exists(config.CONFIG_FILE):
            _create_default_config()
        if not os.path.exists(config.TEACHERS_FILE):
            _create_sample_teachers()
        if not os.path.exists(config.EXAMS_FILE):
            _create_sample_exams()


def _create_default_config():
    """创建默认配置"""
    data = [{
        '配置项': '每个老师每天最多监考次数',
        '值': config.DEFAULT_MAX_EXAMS_PER_DAY
    }, {
        '配置项': '连续监考最多场次',
        '值': config.DEFAULT_MAX_CONSECUTIVE_EXAMS
    }, {
        '配置项': '时间段1',
        '值': config.DEFAULT_TIME_SLOTS[0] if len(config.DEFAULT_TIME_SLOTS) > 0 else ''
    }, {
        '配置项': '时间段2',
        '值': config.DEFAULT_TIME_SLOTS[1] if len(config.DEFAULT_TIME_SLOTS) > 1 else ''
    }, {
        '配置项': '时间段3',
        '值': config.DEFAULT_TIME_SLOTS[2] if len(config.DEFAULT_TIME_SLOTS) > 2 else ''
    }, {
        '配置项': '时间段4',
        '值': config.DEFAULT_TIME_SLOTS[3] if len(config.DEFAULT_TIME_SLOTS) > 3 else ''
    }]
    df = pd.DataFrame(data)
    df.to_excel(config.CONFIG_FILE, index=False)
    print(f"创建默认配置: {config.CONFIG_FILE}")


def load_config() -> Dict[str, Any]:
    """加载配置"""
    try:
        if not os.path.exists(config.CONFIG_FILE):
            _create_default_config()
        
        df = pd.read_excel(config.CONFIG_FILE)
        if df.empty:
            _create_default_config()
            df = pd.read_excel(config.CONFIG_FILE)
        
        config_dict = {}
        for _, row in df.iterrows():
            key = str(row['配置项'])
            value = row['值']
            config_dict[key] = value
        
        return config_dict
    except Exception as e:
        print(f"加载配置失败: {e}")
        return {}


def save_config(config_dict: Dict[str, Any]):
    """保存配置"""
    data = []
    for key, value in config_dict.items():
        data.append({
            '配置项': key,
            '值': value
        })
    df = pd.DataFrame(data)
    df.to_excel(config.CONFIG_FILE, index=False)
    print(f"配置已保存: {config.CONFIG_FILE}")


def _create_sample_teachers():
    """创建示例监考老师数据"""
    sample_teachers = [
        {"工号": "T001", "姓名": "张老师", "职称": "副教授", "联系方式": "13800000001", "所属部门": "计算机学院"},
        {"工号": "T002", "姓名": "李老师", "职称": "讲师", "联系方式": "13800000002", "所属部门": "计算机学院"},
        {"工号": "T003", "姓名": "王老师", "职称": "教授", "联系方式": "13800000003", "所属部门": "数学学院"},
        {"工号": "T004", "姓名": "赵老师", "职称": "讲师", "联系方式": "13800000004", "所属部门": "物理学院"},
        {"工号": "T005", "姓名": "刘老师", "职称": "副教授", "联系方式": "13800000005", "所属部门": "化学学院"},
        {"工号": "T006", "姓名": "陈老师", "职称": "讲师", "联系方式": "13800000006", "所属部门": "计算机学院"},
        {"工号": "T007", "姓名": "周老师", "职称": "教授", "联系方式": "13800000007", "所属部门": "数学学院"},
        {"工号": "T008", "姓名": "吴老师", "职称": "副教授", "联系方式": "13800000008", "所属部门": "物理学院"},
    ]

    df = pd.DataFrame(sample_teachers)
    df.to_excel(config.TEACHERS_FILE, index=False)
    print(f"创建示例监考老师数据: {config.TEACHERS_FILE}")


def _create_sample_exams():
    """创建示例考试数据"""
    sample_exams = [
        {"考试编号": "E001", "考试名称": "期末考试-高等数学", "科目": "高等数学", "日期": "2024-06-15", "时间段": "08:30-10:30", "考场": "A101", "需要监考人数": 2, "考场数": 6},
        {"考试编号": "E002", "考试名称": "期末考试-高等数学", "科目": "高等数学", "日期": "2024-06-15", "时间段": "10:45-12:45", "考场": "A101", "需要监考人数": 2, "考场数": 6},
        {"考试编号": "E003", "考试名称": "期末考试-程序设计", "科目": "程序设计", "日期": "2024-06-15", "时间段": "14:00-16:00", "考场": "B201", "需要监考人数": 2, "考场数": 6},
        {"考试编号": "E004", "考试名称": "期末考试-数据结构", "科目": "数据结构", "日期": "2024-06-16", "时间段": "08:30-10:30", "考场": "A102", "需要监考人数": 2, "考场数": 6},
        {"考试编号": "E005", "考试名称": "期末考试-大学物理", "科目": "大学物理", "日期": "2024-06-16", "时间段": "14:00-16:00", "考场": "C301", "需要监考人数": 2, "考场数": 6},
    ]

    df = pd.DataFrame(sample_exams)
    df.to_excel(config.EXAMS_FILE, index=False)
    print(f"创建示例考试数据: {config.EXAMS_FILE}")


def save_teachers(teachers: List[Teacher]):
    """保存监考老师数据"""
    data = []
    for t in teachers:
        data.append({
            '工号': t.teacher_id,
            '姓名': t.name,
            '职称': t.title,
            '联系方式': t.phone,
            '所属部门': t.department
        })
    df = pd.DataFrame(data)
    df.to_excel(config.TEACHERS_FILE, index=False)
    print(f"保存监考老师数据: {config.TEACHERS_FILE}")


def save_exams(exams: List[Exam]):
    """保存考试数据"""
    data = []
    for e in exams:
        data.append({
            '考试编号': e.exam_id,
            '考试名称': e.exam_name,
            '科目': e.subject,
            '日期': e.date,
            '时间段': e.time_slot,
            '考场': e.room,
            '需要监考人数': e.required_teachers,
            '考场数': e.rooms_count
        })
    df = pd.DataFrame(data)
    df.to_excel(config.EXAMS_FILE, index=False)
    print(f"保存考试数据: {config.EXAMS_FILE}")


def load_teachers() -> List[Teacher]:
    """加载监考老师数据"""
    if not os.path.exists(config.TEACHERS_FILE):
        print(f"警告: 监考老师文件不存在: {config.TEACHERS_FILE}")
        return []

    df = pd.read_excel(config.TEACHERS_FILE)
    teachers = []

    for idx, row in df.iterrows():
        teacher_id_value = row.get('工号', None)
        if pd.isna(teacher_id_value) or str(teacher_id_value) == 'nan':
            teacher_id = f"T{idx+1:03d}"
        else:
            teacher_id = str(teacher_id_value)

        teacher = Teacher(
            teacher_id=teacher_id,
            name=str(row['姓名']),
            title=str(row['职称']),
            phone=str(row['联系方式']),
            department=str(row['所属部门']),
            exam_count=0
        )
        teachers.append(teacher)

    print(f"加载监考老师: {len(teachers)} 人")
    return teachers


def load_exams() -> List[Exam]:
    """加载考试数据"""
    if not os.path.exists(config.EXAMS_FILE):
        print(f"警告: 考试文件不存在: {config.EXAMS_FILE}")
        return []

    df = pd.read_excel(config.EXAMS_FILE)
    exams = []

    for _, row in df.iterrows():
        required_value = row.get('需要监考人数', 2)

        if required_value is None or str(required_value) != str(required_value):
            required_value = 2
        required_teachers = int(required_value)

        # 获取考场数，默认为6
        rooms_count_value = row.get('考场数', 6)
        if rooms_count_value is None or str(rooms_count_value) != str(rooms_count_value):
            rooms_count_value = 6
        rooms_count = int(rooms_count_value)

        exam = Exam(
            exam_id=str(row['考试编号']),
            exam_name=str(row['考试名称']),
            subject=str(row['科目']),
            date=str(row['日期']),
            time_slot=str(row['时间段']),
            room=str(row['考场']),
            required_teachers=required_teachers,
            rooms_count=rooms_count
        )
        exams.append(exam)

    print(f"加载考试: {len(exams)} 场")
    return exams


def export_schedule(schedules: List[Schedule]):
    """导出排班结果到Excel"""
    if not schedules:
        print("警告: 没有排班数据可导出")
        return

    data = []
    for schedule in schedules:
        room = schedule.exam.room if schedule.exam.room else schedule.exam.subject
        
        for i, teacher in enumerate(schedule.teachers):
            row = {
                '日期': schedule.exam.date,
                '时间段': schedule.exam.time_slot,
                '考场': room,
                '考试名称': schedule.exam.exam_name,
                '科目': schedule.exam.subject,
                '需要监考人数': schedule.exam.required_teachers,
                '监考教师': teacher.name,
                '监考教师工号': teacher.teacher_id if teacher.teacher_id else '',
                '职位': i == 0 and '主监考' or '副监考'
            }
            data.append(row)
    
    df = pd.DataFrame(data)
    df.to_excel(config.SCHEDULE_FILE, index=False)
    print(f"排班结果已导出到: {config.SCHEDULE_FILE}")


def export_schedule_by_date(schedules: List[Schedule]):
    """按日期导出排班结果"""
    dates = sorted(set([s.exam.date for s in schedules]))

    for date in dates:
        date_schedules = [s for s in schedules if s.exam.date == date]
        filename = f"{config.DATA_DIR}/schedule_{date}.xlsx"

        data = []
        for schedule in date_schedules:
            teachers_str = "、".join([f"{t.name}({t.teacher_id})" for t in schedule.teachers])
            data.append({
                '时间段': schedule.exam.time_slot,
                '考场': schedule.exam.room,
                '考试名称': schedule.exam.exam_name,
                '科目': schedule.exam.subject,
                '监考老师': teachers_str
            })

        df = pd.DataFrame(data)
        df.to_excel(filename, index=False)
        print(f"日期 {date} 的排班结果已导出到: {filename}")


def print_schedule(schedules: List[Schedule]):
    """打印排班结果"""
    if not schedules:
        print("没有排班数据")
        return

    print("\n" + "=" * 100)
    print(f"{'日期':<12} {'时间段':<15} {'考场':<8} {'考试名称':<20} {'科目':<15} {'监考老师'}")
    print("=" * 100)

    for schedule in schedules:
        teachers_str = "、".join([f"{t.name}({t.teacher_id})" for t in schedule.teachers])
        print(f"{schedule.exam.date:<12} {schedule.exam.time_slot:<15} {schedule.exam.room:<8} "
              f"{schedule.exam.exam_name:<20} {schedule.exam.subject:<15} {teachers_str}")

    print("=" * 100)


def print_statistics(stats: dict):
    """打印统计信息"""
    print("\n" + "=" * 80)
    print("排班统计信息")
    print("=" * 80)
    print(f"总考试数: {stats['total_exams']}")
    print(f"已排班: {stats['scheduled_exams']}")
    print(f"未排班: {stats['unscheduled_exams']}")
    print("\n老师监考统计:")
    print("-" * 80)
    print(f"{'工号':<10} {'姓名':<10} {'监考次数'}")
    print("-" * 80)

    for t_stat in stats['teacher_stats']:
        print(f"{t_stat['teacher_id']:<10} {t_stat['name']:<10} {t_stat['exam_count']}")

    print("=" * 80)
