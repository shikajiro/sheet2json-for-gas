/*
 * シートをjson文字列に変換する。
 * シート間の#id と #refer を一致させることで、
 * 階層構造のjsonも作成できる。
 * @author shikajiro
 */

function onOpen() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [ {name: "sheet2json", functionName: "sheet2json"}];
  ss.addMenu("scripts", menuEntries);
}

/*
 * main
 */
function sheet2json() {
  Logger.log("sheet2json:"+new Date().toLocaleString());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var as = ss.getActiveSheet();
  var id = "";
  var jsondata = generateSheet2Json(as,id, "array");

  Browser.msgBox("json data is: " + jsondata); 
}

/*
 *シートをjsonに変換するkeyとvalueの範囲を指定する。
 */
function generateSheet2Json(sheet, id, type){
  Logger.log("generateSheet2Json start:"+new Date().toLocaleString());
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ranges = ss.getRangeByName(sheet.getName());
  var keys = ranges.getValues()[0];
  var values = ranges.getValues().slice(1);
  var jsons = generateJsons(keys, values, id, type);
  Logger.log("generateSheet2Json end:"+new Date().toLocaleString());
  return jsons;
}

/*
 *配列からjson文字列を作成する
 */
function generateJsons(keys, values, id, type){
  var jsons = "";
  var selection = [];
  //idが指定されている場合、referと一致するレコードだけ抽出する。
  if(id){
    for(var v in values){
      index = keys.indexOf("#refer");
      if(index != -1){
        if(id == values[v][index]){
          selection.push(values[v]);
        }
      }
    }
  //シートを再帰的に参照している場合、最初の呼び出しをreferが無いレコードのみにする。
  }else if(!id && keys.indexOf("#refer") != -1){
    for(var r in values){
        index = keys.indexOf("#refer");
        if(!values[r][index]){
          selection.push(values[r]);
        }
    }
  }else{
    selection = values;
  }

  //  
  for(var i in selection){
    jsons += generateJson(keys, selection[i], id);
    if(i != selection.length-1){
      jsons += ",";
    }
  }
  if(type == "array" && jsons){
    jsons = "[" + jsons + "]";
  }
  return jsons;
}

/*
 *一つの行集合からjson文字列を作成する。
 */
function generateJson(keys, values, id){
  var json = "";
  for(var i in values){
    var value = ""+values[i];
    
    //#idはjsonに出力しない。
    if(keys[i].search("#id") != -1){
      id = value;
      continue;
    }

    //#referはjsonに出力しない。
    if(keys[i].search("#refer") != -1){
      continue;
    }

    //#が含まれるvalueはシート参照先とみなす
    if(value.indexOf("#") != -1){
      var obj_type = (value.search("]") == -1) ? "obj" : "array";
      var slice = 0;
      if(obj_type == "array"){
        slice = 3;
      }else{
        slice = 1;
      }
      var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(value.slice(slice));
      value = generateSheet2Json(sheet,id, obj_type);
    };  
                                   
    json += generateKeyValue(keys[i], value);
    if(i != values.length-1){
      json += ",";
    }
  }
  if(json){
    json = "{" + json + "}";
  }
  return json;
}
  
/*
 *keyとvalueからjson文字列を作成する。
 */
function generateKeyValue(key, value){
  var valueStr = "";
  if(value.indexOf("[") == 0 || value.indexOf("{") == 0){
    valueStr = value;
  }else{
    valueStr = '\"'+ value +'\"';
  }
  var hash = '\"'+ key +'\":'+valueStr; 
  return hash;
} 