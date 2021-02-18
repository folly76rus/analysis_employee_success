# -*- coding: utf-8 -*-
from success_analyzer import SuccessAnalyzer


path_read = "excel/"
path_output = "excel_assessment/"
analyzer = SuccessAnalyzer(path_read=path_read, path_output=path_output)
analyzer.print_united_table()
analyzer.output_to_console()