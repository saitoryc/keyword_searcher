require './excel'
require './file-win32'
require 'open-uri'

filename = FileSystemObject.instance.getAbsolutePathName(ARGV[0])
excel = Excel.new
workbook = excel.workbooks.open({'filename'=> filename, 'readOnly' => false})
ws = workbook.sheets[1]
current_row=5

while ws.rows[current_row].columns[4].value != "" do
	url = ws.rows[current_row].columns[4].value
	searchWord = ws.rows[current_row].columns[5].value

	sio = OpenURI.open_uri(url)
	check = sio.read.include?(searchWord)
	if check
		ws.rows[current_row].columns[8].value = "OK"
		ws.rows[current_row].Interior.ColorIndex = 43

		lines = searchWord.split("¥n")
		lineCnt = 1
		for line in lines do
			if (line.include?(searchWord))
				ws.rows[current_row].columns[9].value = lineCnt
				ws.rows[current_row].columns[10].value = line
				break
			end
			lineCnt = lineCnt + 1
		end
	else
		ws.rows[current_row].columns[8].value = "NG"
		ws.rows[current_row].Interior.ColorIndex = 3
	end

	current_row = current_row + 1
end


workbook.save
