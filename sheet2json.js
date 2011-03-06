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
  
  init();
  var jsondata = generateSheet2Json(as,id, "array");
  destroy();
  
  Browser.msgBox("json data is: " + jsondata);
  
  /*
  * 初期化。名前付けされた範囲を設定する。
  */
  function init(){
    var sheets = ss.getSheets();
    for(var i in sheets){
      var sheet = sheets[i];
      //既存の名前付けされた範囲を削除
      var name = sheet.getName();
      if(ss.getRangeByName(sheet.getName())){
        ss.removeNamedRange(sheet.getName());
      }
      ss.setNamedRange(sheet.getName(), getNameRanges(sheet));
    }
    
    /*
    * keyとvalueの範囲を指定する。
    */
    function getNameRanges(sheet){
      
      //とりあえずシートを範囲選択する。10000行、255列をMAXとする。
      var values = sheet.getRange(1,1,10000,255).getValues();
      
      //セルの値が無くなるまで横にずれて値があるか探す。
      for(var colIndex in values[0]){
        var key = values[0][colIndex];
        if(!key){
          break;
            }
      }
      //セルの値が無くなるまで下にずれて値があるか探す。
      for(var rowIndex in values){
        var value = values[rowIndex][0];
        if(!value){
          break;
            }
      }
      var valueRange = sheet.getRange(1, 1, parseInt(rowIndex)+1, parseInt(colIndex)+1);
      
      return valueRange;
    }
  };
  
  
  /*
  * 終了処理。
  * 名前付けした範囲をクリアする。
  */
  function destroy(){
    var sheets = ss.getSheets();
    for(var i in sheets){
      var sheet = sheets[i];
      //既存の名前付けされた範囲を削除
      if(ss.getRangeByName(sheet.getName())){
        ss.removeNamedRange(sheet.getName());
      }
    }
  };
  
  
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
  }
  
}

