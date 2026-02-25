# 主程序
# =====

from models import Teacher, Exam
from scheduler import ExamScheduler
from utils import (
    init_data_dir, load_teachers, load_exams,
    export_schedule, export_schedule_by_date,
    print_schedule, print_statistics
)


def main():
    """主函数"""
    print("=" * 80)
    print("监考老师排班系统")
    print("=" * 80)

    # 初始化数据目录和示例数据
    init_data_dir()

    # 加载数据
    print("\n正在加载数据...")
    teachers = load_teachers()
    exams = load_exams()

    if not teachers or not exams:
        print("\n错误: 没有足够的数据进行排班")
        print("请检查 data/teachers.xlsx 和 data/exams.xlsx 文件")
        return

    # 创建排班器
    scheduler = ExamScheduler(teachers, exams)

    while True:
        print("\n" + "=" * 80)
        print("请选择操作:")
        print("=" * 80)
        print("1. 执行自动排班")
        print("2. 查看排班结果")
        print("3. 查看统计信息")
        print("4. 导出排班结果 (Excel)")
        print("5. 按日期导出排班结果")
        print("6. 查看监考老师列表")
        print("7. 查看考试列表")
        print("8. 重新加载数据")
        print("0. 退出")
        print("=" * 80)

        choice = input("\n请输入选项 (0-8): ").strip()

        if choice == '1':
            execute_schedule(scheduler)
        elif choice == '2':
            show_schedule(scheduler)
        elif choice == '3':
            show_statistics(scheduler)
        elif choice == '4':
            export_schedule_excel(scheduler)
        elif choice == '5':
            export_by_date(scheduler)
        elif choice == '6':
            show_teachers(teachers)
        elif choice == '7':
            show_exams(exams)
        elif choice == '8':
            reload_data()
            teachers = load_teachers()
            exams = load_exams()
            scheduler = ExamScheduler(teachers, exams)
        elif choice == '0':
            print("\n感谢使用，再见！")
            break
        else:
            print("\n无效的选项，请重新选择")


def execute_schedule(scheduler: ExamScheduler):
    """执行排班"""
    print("\n开始执行自动排班...")
    schedules = scheduler.schedule()

    if schedules:
        print(f"\n排班成功！共安排了 {len(schedules)} 场考试")
        print_schedule(schedules)
    else:
        print("\n排班失败，请检查数据是否正确")


def show_schedule(scheduler: ExamScheduler):
    """查看排班结果"""
    if not scheduler.final_schedules:
        print("\n暂无排班数据，请先执行排班")
        return

    print("\n请选择查看方式:")
    print("1. 查看全部排班")
    print("2. 按日期查看")
    print("3. 按老师查看")

    choice = input("请输入选项 (1-3): ").strip()

    if choice == '1':
        print_schedule(scheduler.final_schedules)
    elif choice == '2':
        view_by_date(scheduler)
    elif choice == '3':
        view_by_teacher(scheduler)
    else:
        print("\n无效的选项")


def view_by_date(scheduler: ExamScheduler):
    """按日期查看排班"""
    dates = sorted(set([s.exam.date for s in scheduler.final_schedules]))

    if not dates:
        print("\n暂无排班数据")
        return

    print("\n可用日期:")
    for i, date in enumerate(dates, 1):
        print(f"{i}. {date}")

    try:
        index = int(input("\n请输入日期序号: ").strip()) - 1
        if 0 <= index < len(dates):
            date_schedules = scheduler.get_schedule_by_date(dates[index])
            print(f"\n{dates[index]} 的排班情况:")
            print_schedule(date_schedules)
        else:
            print("\n无效的日期序号")
    except ValueError:
        print("\n请输入有效的数字")


def view_by_teacher(scheduler: ExamScheduler):
    """按老师查看排班"""
    print("\n请输入老师工号: ", end="")
    teacher_id = input().strip()

    teacher_schedules = scheduler.get_schedule_by_teacher(teacher_id)

    if not teacher_schedules:
        print(f"\n未找到工号为 {teacher_id} 的老师排班信息")
        return

    print(f"\n工号 {teacher_id} 的排班情况 (共 {len(teacher_schedules)} 场):")
    print_schedule(teacher_schedules)


def show_statistics(scheduler: ExamScheduler):
    """查看统计信息"""
    if not scheduler.final_schedules:
        print("\n暂无排班数据，请先执行排班")
        return

    stats = scheduler.get_statistics()
    print_statistics(stats)


def export_schedule_excel(scheduler: ExamScheduler):
    """导出排班结果"""
    if not scheduler.final_schedules:
        print("\n暂无排班数据，请先执行排班")
        return

    export_schedule(scheduler.final_schedules)


def export_by_date(scheduler: ExamScheduler):
    """按日期导出排班"""
    if not scheduler.final_schedules:
        print("\n暂无排班数据，请先执行排班")
        return

    export_schedule_by_date(scheduler.final_schedules)


def show_teachers(teachers: list):
    """显示监考老师列表"""
    if not teachers:
        print("\n暂无监考老师数据")
        return

    print("\n" + "=" * 100)
    print(f"{'工号':<10} {'姓名':<10} {'职称':<10} {'联系方式':<15} {'所属部门':<20}")
    print("=" * 100)

    for teacher in teachers:
        print(f"{teacher.teacher_id:<10} {teacher.name:<10} {teacher.title:<10} "
              f"{teacher.phone:<15} {teacher.department:<20}")

    print("=" * 100)
    print(f"总计: {len(teachers)} 人")


def show_exams(exams: list):
    """显示考试列表"""
    if not exams:
        print("\n暂无考试数据")
        return

    print("\n" + "=" * 100)
    print(f"{'考试编号':<12} {'考试名称':<20} {'科目':<15} {'日期':<12} {'时间段':<15} {'考场':<8} {'需要人数':<8}")
    print("=" * 100)

    for exam in exams:
        print(f"{exam.exam_id:<12} {exam.exam_name:<20} {exam.subject:<15} "
              f"{exam.date:<12} {exam.time_slot:<15} {exam.room:<8} {exam.required_teachers:<8}")

    print("=" * 100)
    print(f"总计: {len(exams)} 场考试")


def reload_data():
    """重新加载数据"""
    print("\n数据已重新加载")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\n\n程序已中断")
    except Exception as e:
        print(f"\n发生错误: {e}")
        import traceback
        traceback.print_exc()
