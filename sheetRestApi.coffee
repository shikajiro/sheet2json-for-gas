createJson = (appId, sheetName)->
	spreadsheet = SpreadsheetApp.openById appId
	sheet = spreadsheet.getSheetByName sheetName
	dataRange = sheet.getDataRange().getValues()

	#１行目のカラム名を取得
	culumnName = (column for column in dataRange[0])

	#２行目以降のデータが入ったカラムを取得
	table = []
	for row, r in dataRange
		record = {}
		for cell, c in row[1..row.length]
			record[culumnName[c]] = cell
		table.push record

	#jsonに変換して返す
	json = Utilities.jsonStringify table
	Logger.log json
	return json

getJson =(req)->
	#spreadsheetのid
	appId = req.parameters.app_id
	# シートの名前
	sheetName = req.parameters.sheet_name

	output = ContentService.createTextOutput()
	output.setMimeType ContentService.MimeType.JSON
	output.setContent createJson(appId, sheetName)
	return output

postJson = (req)->
	Logger.log req

`function onOpen() { createJson(); }`
`function doGet(req) { Logger.log(req); return getJson(req); }`
`function doPost(req) { Logger.log(req); return postJson(req); }`