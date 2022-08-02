import datetime
import pprint


from effectiveness_report_compliler import EffectivenessReportCompiler
general_directory_path = input('Provide address for general directory\n')

if __name__ == '__main__':
    today = datetime.datetime.weekday(datetime.datetime.today())
    last_monday = datetime.datetime.today() - datetime.timedelta(weeks=2, days=today)
    last_sunday = datetime.datetime.today() - datetime.timedelta(weeks=1, days=today + 1)
    # last_monday = datetime.datetime(year=2022, month=6, day=1)
    # last_sunday = datetime.datetime(year=2022, month=6, day=30)

    report_compile = EffectivenessReportCompiler(start_date=last_monday, end_date=last_sunday,
                                                 general_directory_path=general_directory_path)
    report_compile.act()
    pprint.pprint(report_compile.data_for_eu_total_compiler)
