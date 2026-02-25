"""
数据模型
"""

from dataclasses import dataclass
from typing import List, Optional
from datetime import datetime


@dataclass
class Teacher:
    """监考老师"""
    teacher_id: str  # 工号
    name: str  # 姓名
    title: str  # 职称
    phone: str  # 联系方式
    department: str  # 所属部门
    exam_count: int = 0  # 已监考次数

    def __hash__(self):
        return hash(self.teacher_id)

    def __eq__(self, other):
        if not isinstance(other, Teacher):
            return False
        return self.teacher_id == other.teacher_id


@dataclass
class Exam:
    """考试信息"""
    exam_id: str  # 考试编号
    exam_name: str  # 考试名称
    subject: str  # 考试科目
    date: str  # 考试日期 (YYYY-MM-DD)
    time_slot: str  # 时间段
    room: str  # 考场号
    required_teachers: int = 2  # 每个考场需要的监考老师数
    rooms_count: int = 6  # 该考试需要分配的考场数

    def __hash__(self):
        return hash(self.exam_id)

    def __eq__(self, other):
        if not isinstance(other, Exam):
            return False
        return self.exam_id == other.exam_id


@dataclass
class Schedule:
    """排班信息"""
    exam: Exam  # 考试信息
    teachers: List[Teacher]  # 监考老师列表

    def __str__(self):
        teachers_str = "、".join([t.name for t in self.teachers])
        return f"{self.exam.date} {self.exam.time_slot} {self.exam.room} {self.exam.subject} - 监考老师: {teachers_str}"


@dataclass
class TeacherSchedule:
    """老师排班表"""
    teacher: Teacher
    schedules: List[Schedule]

    def has_conflict(self, exam: Exam) -> bool:
        """检查时间冲突"""
        for schedule in self.schedules:
            if (schedule.exam.date == exam.date and
                schedule.exam.time_slot == exam.time_slot):
                return True
        return False

    def has_conflict_by_time(self, date: str, time_slot: str) -> bool:
        """检查指定时间是否有冲突"""
        for schedule in self.schedules:
            if schedule.exam.date == date and schedule.exam.time_slot == time_slot:
                return True
        return False

    def get_daily_exam_count(self, date: str) -> int:
        """获取当天监考次数"""
        return len([s for s in self.schedules if s.exam.date == date])

    def get_consecutive_count(self, exam: Exam) -> int:
        """获取连续监考次数"""
        sorted_schedules = sorted(self.schedules, key=lambda x: x.exam.date)
        count = 0
        for schedule in sorted_schedules:
            if schedule.exam.date == exam.date:
                count += 1
            else:
                count = 0
        return count
