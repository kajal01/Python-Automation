

#first save the result to the CSV file
def SaveTestResultToCSV ( tab_id, row_no,test_result,remarks) : 
	print("Saving the Test case Document Result to csv file..."+ tab_id + row_no + test_result + remarks)

#then Open the CSV file and dump data to appropriate columns form csv file
def SaveTestResultToExcel ( csv_file) : 
	print("Saving the Test case Document Result to Excel file..."+ csv_file)
	