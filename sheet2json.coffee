###
* シートをjson文字列に変換する。
* シート間の#id と #refer を一致させることで、
* 階層構造のjsonも作成できる。
* @author shikajiro
###

#TODO 日本語除去


log = Logger.log

logTimeMesure = (methodName)->
  log (methodName+":"+new Date().toLocaleString())

onOpen = () ->
  ss = SpreadsheetApp.getActiveSpreadsheet()
  menuEntries = [ name: "sheet2json", functionName: "sheet2json"]
  ss.addMenu("scripts", menuEntries)

class Sheet2json
  constructor: (ss)->

    ###
     # 初期化。名前付けされた範囲を設定する。
    ###
    _init = ->
      for sheet in ss.getSheets()
        #既存の名前付けされた範囲を削除
        name = sheet.getName()
        ss.removeNamedRange(name) if ss.getRangeByName(name)
        ss.setNamedRange(name, _getNameRanges(sheet))
      
    ###
    * keyとvalueの範囲を指定する。
    ###
    _getNameRanges = (sheet)->
      
      #とりあえずシートを範囲選択する。10000行、255列をMAXとする。
      values = sheet.getRange(1,1,10000,255).getValues()
      
      #セルの値が無くなるまで横にずれて値があるか探す。
      break for key, rowIndex in values[0] when !key
      
      #セルの値が無くなるまで下にずれて値があるか探す。
      break for value, colIndex in values when !value[0]
      
      sheet.getRange 1, 1, rowIndex-1, colIndex-1

    ###
    * 終了処理。
    * 名前付けした範囲をクリアする。
    ###
    _destroy = ->
      for sheet in ss.getSheets()
        #既存の名前付けされた範囲を削除
        name = sheet.getName()
        ss.removeNamedRange name if ss.getRangeByName(name)

    ###
    *シートをjsonに変換するkeyとvalueの範囲を指定する。
    ###
    _generateSheet2Json = (sheet, id, type)->
      logTimeMesure "generateSheet2Json start"
      # ss = SpreadsheetApp.getActiveSpreadsheet()
      ranges = ss.getRangeByName(sheet.getName())
      keys = ranges.getValues()[0]
      values = ranges.getValues().slice(1)

      datas = []
      for record in values
        data={}
        data[key] = record[i] for key, i in keys
        datas.push data
      log datas
      # selection = _generateJsons(keys, values, id, type)
      logTimeMesure "generateSheet2Json end"
      json = Utilities.jsonStringify(datas)
      json

    ###
    *配列からjson文字列を作成する
    ###
    _generateJsons = (keys, values, id, type)->
      selection = []
      refer = keys.indexOf("#refer")
      #idが指定されている場合、referと一致するレコードだけ抽出する。
      if id
        for rec in values
          selection.push(rec) if refer != -1 and id is rec[index]
        
      #シートを再帰的に参照している場合、最初の呼び出しをreferが無いレコードのみにする。
      else if !id and refer != -1
        for rec in values
          selection.push(rec) if !rec[refer]
        
      else
        selection = values
      
        selection

      # jsons = ""
      # for value, i in selection
      #   jsons += _generateJson keys, value, id
      #   jsons += "," if(i != selection.length - 1)
      
      # jsons = "[" + jsons + "]" if jsons and type == "array"
      # jsons

    ###
    *一つの行集合からjson文字列を作成する。
    ###
    _generateJson = (keys, values, id)->
      json = ""
      for value, i in values
        value = ""+value 
        key = keys[i]
        
        ##idはjsonに出力しない。
        if(key.search("#id") != -1)
          id = value
          continue
        
        ##referはjsonに出力しない。
        if(key.search("#refer") != -1)
          continue
        
        ##が含まれるvalueはシート参照先とみなす
        if(value.indexOf("#") != -1)
          obj_type = (value.search("]") == -1) ? "obj" : "array"
          slice = 0
          if(obj_type == "array")
            slice = 3
          else
            slice = 1
          
          sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(value.slice(slice))
          selection = _generateSheet2Json(sheet,id, obj_type)
        
        # json += _generateKeyValue(keys[i], value)
        # json += "," if i != values.length-1
      
      # json = "{" + json + "}" if(json)
      # json

    ###
    *keyとvalueからjson文字列を作成する。
    ###
    # _generateKeyValue = (key, value)->
    #   valueStr = ""
    #   if value.indexOf("[") == 0 || value.indexOf("{") == 0
    #     valueStr = value
    #   else
    #     valueStr = '\"'+value+'\"'
      
    #   hash = '\"'+key+'\":'+valueStr 
    #   hash

    ###
    #  main
    ###
    @sheet2json = ->
      logTimeMesure "sheet2json"
      ss = SpreadsheetApp.getActiveSpreadsheet()
      as = ss.getActiveSheet()
      id = ""
      
      _init()
      jsondata = _generateSheet2Json(as,id, "array")
      _destroy()
      
      Browser.msgBox(jsondata)
      return
    
`function sheet2json() { new Sheet2json().sheet2json(); }`
`function onOpen() { new Sheet2json().onOpen(); }`
