import os


class LogFile:
    def __init__(self, file_path: str) -> None:
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.scenario_name = self.file_name.split("_")[3]
        self.sub_scenario_name = None

        if "(" in self.file_name:
            first_index = self.file_name.index("(")
            last_index = self.file_name.index(")")
            self.sub_scenario_name = self.file_name[first_index + 1 : last_index]

    def __str__(self) -> str:
        return self.file_name


logfile = LogFile(
    r"C:\Users\shpar\Documents\Log_20211001131853_Unknown Road_시나리오1_1시간_1_0_0_0.csv"
)

print(logfile.sub_scenario_name)
