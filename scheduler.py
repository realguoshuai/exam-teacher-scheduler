"""
排班算法
"""

from models import Teacher, Exam, Schedule, TeacherSchedule
from typing import List, Dict, Set, Tuple, Optional
import random


class ExamScheduler:
    """考试排班系统"""

    def __init__(self, teachers: List[Teacher], exams: List[Exam], config: Optional[Dict] = None):
        self.teachers = teachers
        self.exams = exams
        self.config = config or {}

        self.teacher_schedules: Dict[str, TeacherSchedule] = {
            t.teacher_id: TeacherSchedule(teacher=t, schedules=[])
            for t in teachers
        }
        self.final_schedules: List[Schedule] = []

    def schedule(self) -> List[Schedule]:
        """执行排班"""
        for teacher in self.teachers:
            teacher.exam_count = 0
        
        self.final_schedules = []
        
        unique_exams = self._deduplicate_exams(self.exams)
        print(f"去重后考试数: {len(unique_exams)} 场")

        exams_by_time = {}
        for exam in unique_exams:
            key = (exam.date, exam.time_slot)
            if key not in exams_by_time:
                exams_by_time[key] = []
            exams_by_time[key].append(exam)

        # 按时间排序处理
        sorted_times = sorted(exams_by_time.keys())
        
        for date, time_slot in sorted_times:
            exams = exams_by_time[(date, time_slot)]
            print(f"\n时间段 {date} {time_slot}:")
            self._schedule_exams_at_time(date, time_slot, exams)
            
            # 每次分配后检查全局平衡
            self._check_and_balance()

        print(f"\n总共生成排班: {len(self.final_schedules)} 条记录")
        return self.final_schedules
    
    def _check_and_balance(self):
        """检查并平衡老师排班次数，确保差距不超过2"""
        if not self.teachers:
            return
            
        counts = [t.exam_count for t in self.teachers]
        max_count = max(counts)
        min_count = min(counts)
        
        # 如果最大差距超过2，打印警告
        if max_count - min_count > 2:
            print(f"  [平衡检查] 警告：老师排班次数差距过大（最大{max_count}, 最小{min_count}）")

    def _deduplicate_exams(self, exams: List[Exam]) -> List[Exam]:
        seen = set()
        unique = []
        for exam in exams:
            key = (exam.exam_name, exam.date, exam.time_slot, exam.subject)
            if key not in seen:
                seen.add(key)
                unique.append(exam)
        return unique

    def _schedule_exams_at_time(self, date: str, time_slot: str, exams: List[Exam]):
        """为同一时间的多个考试安排老师"""
        subjects = list(set(e.subject for e in exams))
        print(f"  科目: {subjects} (共 {len(subjects)} 个科目)")

        # 初始化该时间段需要分配的考场列表
        rooms_to_assign = []
        for exam in exams:
            subject = exam.subject
            required = exam.required_teachers  # 每个考场需要的老师数（默认2）
            rooms_count = exam.rooms_count  # 该考试的考场数
            for room_num in range(1, rooms_count + 1):
                room = f"{subject}考场{room_num}"
                rooms_to_assign.append({
                    'exam': exam,
                    'subject': subject,
                    'room': room,
                    'required': required,
                    'exam_id': exam.exam_id
                })
        
        total_teachers_needed = sum(r['required'] for r in rooms_to_assign)
        print(f"  需要总考场数: {len(rooms_to_assign)}")
        print(f"  需要总老师数: {total_teachers_needed}")
        
        # 获取在该时间段没有冲突的可用老师
        teachers_with_counts = []
        for teacher in self.teachers:
            teacher_schedule = self.teacher_schedules[teacher.teacher_id]
            if teacher_schedule.has_conflict_by_time(date, time_slot):
                continue
            teachers_with_counts.append(teacher)

        print(f"  可用老师数: {len(teachers_with_counts)}")
        
        # 按监考次数分组排序（公平原则：优先选次数少的）
        # 在相同次数的老师中随机排序，增加公平性
        teachers_with_counts.sort(key=lambda t: (t.exam_count, random.random()))
        
        # 为每个考场分配老师
        assigned_teachers = set()  # 记录该时段已分配的老师ID
        for room_info in rooms_to_assign:
            required = room_info['required']
            teachers_for_room = []

            # 找(required)个未分配的老师，优先选择exam_count少的
            # 每次选择后重新排序，确保始终选择当前次数最少的
            available_teachers = [t for t in teachers_with_counts if t.teacher_id not in assigned_teachers]
            available_teachers.sort(key=lambda t: (t.exam_count, random.random()))
            
            for teacher in available_teachers[:required]:
                teachers_for_room.append(teacher)
                assigned_teachers.add(teacher.teacher_id)
            
            if teachers_for_room:
                exam_copy = Exam(
                    exam_id=f"{room_info['exam_id']}_{room_info['room']}",
                    exam_name=room_info['exam'].exam_name,
                    subject=room_info['subject'],
                    date=date,
                    time_slot=time_slot,
                    room=room_info['room'],
                    required_teachers=room_info['required']
                )
                schedule = Schedule(exam=exam_copy, teachers=teachers_for_room)
                self.final_schedules.append(schedule)
                for teacher in teachers_for_room:
                    self.teacher_schedules[teacher.teacher_id].schedules.append(schedule)
                    teacher.exam_count += 1
                print(f"    {room_info['room']}: {len(teachers_for_room)} 位老师 - {[t.name for t in teachers_for_room]}")
            else:
                print(f"    警告: {room_info['room']} 没有可用老师")

        print(f"  该时间段已分配不同老师数: {len(assigned_teachers)}/{len(teachers_with_counts)}")

    def get_statistics(self) -> Dict:
        """获取统计信息"""
        # 计算总考试数：去重后的每个考试的考场数之和
        unique_exams = self._deduplicate_exams(self.exams)
        total_exams = sum(exam.rooms_count for exam in unique_exams)
        total_schedules = len(self.final_schedules)

        teacher_exam_count = {}
        teacher_exams = {}
        for schedule in self.final_schedules:
            for teacher in schedule.teachers:
                if teacher.teacher_id not in teacher_exam_count:
                    teacher_exam_count[teacher.teacher_id] = 0
                    teacher_exams[teacher.teacher_id] = []
                teacher_exam_count[teacher.teacher_id] += 1
                teacher_exams[teacher.teacher_id].append({
                    'exam_name': schedule.exam.exam_name,
                    'subject': schedule.exam.subject,
                    'date': schedule.exam.date,
                    'time_slot': schedule.exam.time_slot,
                    'room': schedule.exam.room
                })

        teacher_stats = []
        for teacher_id, ts in self.teacher_schedules.items():
            exam_count = teacher_exam_count.get(teacher_id, 0) or 0
            exams = teacher_exams.get(teacher_id, [])
            teacher_stats.append({
                'teacher_id': teacher_id,
                'name': ts.teacher.name,
                'exam_count': exam_count,
                'exams': exams
            })

        teacher_stats = sorted(teacher_stats, key=lambda x: x['exam_count'], reverse=True)

        date_stats = {}
        for schedule in self.final_schedules:
            date = schedule.exam.date
            if date not in date_stats:
                date_stats[date] = {'count': 0, 'teachers': set()}
            date_stats[date]['count'] += 1
            for teacher in schedule.teachers:
                date_stats[date]['teachers'].add(teacher.name)

        for date in date_stats:
            date_stats[date]['teachers'] = list(date_stats[date]['teachers'])

        return {
            'total_exams': total_exams,
            'scheduled_exams': total_schedules,
            'unscheduled_exams': total_exams - total_schedules,
            'teacher_stats': teacher_stats,
            'date_stats': date_stats
        }

    def get_schedule_by_date(self, date: str) -> List[Schedule]:
        return [s for s in self.final_schedules if s.exam.date == date]

    def get_schedule_by_teacher(self, teacher_id: str) -> List[Schedule]:
        if teacher_id in self.teacher_schedules:
            return self.teacher_schedules[teacher_id].schedules
        return []
