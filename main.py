from reportAutomation import ReportAutomation

def main():
	
	automate = ReportAutomation()

	automate.config_setup()
	automate.generate_dates()
	automate.format_dates()
	automate.setup_excel()
	automate.populate_cells()
	automate.save_workbook()

if __name__ == "__main__":
	main()